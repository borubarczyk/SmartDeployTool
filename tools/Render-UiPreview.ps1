# Renderuje podgląd wszystkich okien Smart Tool for Deployment do plików PNG - bez klikania.
#
# Skrypt wczytuje narzędzie w trybie testowym (bez logowania i bez pokazywania okna głównego),
# a następnie po kolei otwiera każde okno (główne, komunikaty, Ustawienia, logi, deinstalator...)
# z prawdziwymi danymi. Globalny handler zdarzenia Loaded renderuje każde otwarte okno do PNG
# i od razu je zamyka, więc wywołania ShowDialog() nie czekają na użytkownika.
#
# Użycie (Windows PowerShell 5.1, wątek STA):
#   powershell -NoProfile -STA -ExecutionPolicy Bypass -File tools\Render-UiPreview.ps1 -Theme Dark
#   powershell -NoProfile -STA -ExecutionPolicy Bypass -File tools\Render-UiPreview.ps1 -Theme Light
# Wynik: ui-preview\<motyw>\NN-nazwa.png (uruchamiane też automatycznie w GitHub Actions,
# workflow "UI Preview" - pliki są do pobrania jako artefakt).
#
# Oprócz zwykłych zrzutów: okno główne w trakcie wdrożenia, w pauzie i z powiadomieniami, okno główne
# i Ustawienia w minimalnym rozmiarze oraz na ekranie 1366x768 przy skalowaniu 150% (911 x 472 DIP).
# Kontrole (lista w konsoli i w plikach obok PNG):
#   CLIP   ucięty tekst                                          -> uciety-tekst.txt
#   STALE  kolor z poprzedniego motywu po przełączeniu w locie   -> uwagi-motyw.txt
#   GLYPH  znak ikony (U+E000-U+F8FF) w zwykłej czcionce         -> uwagi-motyw.txt
#   EMOJI  emoji w tekście okna                                  -> uwagi-motyw.txt
#   CONTRAST  tekst za słabo odcina się od tła (< 3:1)           -> uwagi-motyw.txt
param(
    [ValidateSet('Dark', 'Light')][string]$Theme = 'Dark',
    [string]$OutDir,
    # Wypisuje do konsoli mapę układu (typ, nazwa, tekst, pozycja i rozmiar kontrolek) każdego okna.
    [switch]$DumpLayout,
    # Kończy skrypt kodem 1, jeśli w którymkolwiek oknie wykryto ucięty tekst.
    [switch]$FailOnClipping,
    # Kończy skrypt kodem 1 przy jakiejkolwiek uwadze (CLIP, STALE, GLYPH, EMOJI, CONTRAST) albo błędzie renderowania okna (ERR); obejmuje -FailOnClipping.
    [switch]$FailOnFindings
)
if ($FailOnFindings) { $FailOnClipping = [switch]$true }

$repoRoot = Split-Path -Parent $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($OutDir)) { $OutDir = Join-Path $repoRoot ("ui-preview\" + $Theme.ToLower()) }
New-Item -ItemType Directory -Path $OutDir -Force | Out-Null
$global:UiShotDir = (Resolve-Path $OutDir).Path
$global:UiShotIndex = 0
$global:UiShotName = 'okno'
$global:UiPending = New-Object System.Collections.ArrayList
$global:UiClipping = New-Object System.Collections.ArrayList
$global:UiFindings = New-Object System.Collections.ArrayList
$global:UiErrors = New-Object System.Collections.ArrayList
$global:UiDumpLayout = [bool]$DumpLayout
# Dodatkowe stany okna (klucz = nazwa zrzutu): Apply ustawia stan, Revert przywraca. Okna zamkniętego
# nie da się pokazać drugi raz, więc stany zrzucamy na tym samym, wciąż otwartym oknie (lista niżej).
$global:UiExtraViews = @{}

# Tryb testowy: narzędzie nie pokaże okien logowania/powitania ani nie wywoła ShowDialog okna głównego.
$global:PesterTesting = $true
. (Join-Path $repoRoot 'SmartToolforDeployment.ps1')

function global:Save-UiElementPng {
    param([System.Windows.FrameworkElement]$Element, [System.Windows.Media.Brush]$Background, [string]$Name)
    $Element.UpdateLayout()
    $w = [int][Math]::Ceiling($Element.ActualWidth)
    $h = [int][Math]::Ceiling($Element.ActualHeight)
    if ($w -le 0 -or $h -le 0) { Write-Host "SKIP $Name (rozmiar $w x $h)"; return }
    $rtb = New-Object System.Windows.Media.Imaging.RenderTargetBitmap -ArgumentList $w, $h, 96, 96, ([System.Windows.Media.PixelFormats]::Pbgra32)
    if ($null -ne $Background) {
        $dv = New-Object System.Windows.Media.DrawingVisual
        $dc = $dv.RenderOpen()
        $dc.DrawRectangle($Background, $null, (New-Object System.Windows.Rect 0, 0, $w, $h))
        $dc.Close()
        $rtb.Render($dv)
    }
    $rtb.Render($Element)
    $enc = New-Object System.Windows.Media.Imaging.PngBitmapEncoder
    $enc.Frames.Add([System.Windows.Media.Imaging.BitmapFrame]::Create($rtb))
    $global:UiShotIndex++
    $safe = ($Name -replace '[^\w\-]', '_')
    $file = Join-Path $global:UiShotDir ('{0:D2}-{1}.png' -f $global:UiShotIndex, $safe)
    $fs = [System.IO.File]::Create($file)
    try { $enc.Save($fs) } finally { $fs.Close() }
    Write-Host ("OK   {0} ({1} x {2})" -f (Split-Path $file -Leaf), $w, $h)
}

# Wykrywanie uciętego tekstu: dla każdego widocznego TextBlocka liczymy naturalny rozmiar tekstu
# (FormattedText) i porównujemy z miejscem, które dostał - zarówno sam TextBlock, jak i każdy jego
# rodzic, który przycina zawartość (np. przycisk ze sztywną wysokością albo zbyt wąska kolumna).
# Przycinanie przez ScrollViewer (przewijanie) jest zamierzone i pomijane.
function global:Get-UiElementLabel {
    param($Element)
    $label = $Element.GetType().Name
    if ($Element -is [System.Windows.FrameworkElement] -and -not [string]::IsNullOrWhiteSpace($Element.Name)) { $label += "#$($Element.Name)" }
    return $label
}

function global:Find-UiClipping {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $results = New-Object System.Collections.Generic.List[string]
    $root = [System.Windows.Media.VisualTreeHelper]::GetChild($Win, 0)
    $pixelsPerDip = 1.0
    try { $pixelsPerDip = [System.Windows.Media.VisualTreeHelper]::GetDpi($Win).PixelsPerDip } catch {}
    $stack = New-Object System.Collections.Stack
    $stack.Push($root)
    while ($stack.Count -gt 0) {
        $v = $stack.Pop()
        if ($v -is [System.Windows.UIElement] -and -not $v.IsVisible) { continue }
        for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v); $i++) {
            $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i))
        }
        if ($v -isnot [System.Windows.Controls.TextBlock]) { continue }
        $tb = $v
        if ([string]::IsNullOrWhiteSpace($tb.Text) -or $tb.ActualWidth -le 0) { continue }

        # Najbliższy "właściciel" tekstu do opisu (przycisk, pole, etykieta...).
        $owner = $null
        $p = [System.Windows.Media.VisualTreeHelper]::GetParent($tb)
        while ($null -ne $p -and $null -eq $owner) {
            if ($p -is [System.Windows.Controls.Control]) { $owner = $p }
            $p = [System.Windows.Media.VisualTreeHelper]::GetParent($p)
        }
        $ownerLabel = if ($null -ne $owner) { Get-UiElementLabel $owner } else { 'okno' }
        $textShort = ($tb.Text -replace '\s+', ' ').Trim()
        if ($textShort.Length -gt 60) { $textShort = $textShort.Substring(0, 60) + '…' }

        # 1) Czy tekst mieści się w samym TextBlocku.
        $typeface = New-Object System.Windows.Media.Typeface -ArgumentList $tb.FontFamily, $tb.FontStyle, $tb.FontWeight, $tb.FontStretch
        # Pomiar w trybie, w którym TextBlock się rysuje (okna narzędzia: Display - szerokości liter
        # zaokrąglane do pikseli); domyślny Ideal różnił się o kilka px i dawał fałszywe alarmy.
        # Konstruktor z 9 argumentami (NumberSubstitution = $null) jest w .NET od 4.6.2.
        $mode = [System.Windows.Media.TextOptions]::GetTextFormattingMode($tb)
        $ft = New-Object System.Windows.Media.FormattedText -ArgumentList $tb.Text, ([System.Globalization.CultureInfo]::CurrentUICulture), $tb.FlowDirection, $typeface, $tb.FontSize, ([System.Windows.Media.Brushes]::Black), $null, $mode, $pixelsPerDip
        if (-not [double]::IsNaN($tb.LineHeight)) { $ft.LineHeight = $tb.LineHeight }
        $innerW = $tb.ActualWidth - $tb.Padding.Left - $tb.Padding.Right
        $innerH = $tb.ActualHeight - $tb.Padding.Top - $tb.Padding.Bottom
        if ($tb.TextWrapping -ne [System.Windows.TextWrapping]::NoWrap) {
            # Zawijanie liczymy tylko, gdy tekst wyraźnie nie mieści się w jednej linii - FormattedText
            # mierzy emoji (czcionka zastępcza) o 1-3 px inaczej niż TextBlock, co dawało fałszywe alarmy.
            if ($innerW -gt 0 -and $ft.WidthIncludingTrailingWhitespace -gt $innerW + 4) { $ft.MaxTextWidth = $innerW }
            if ($ft.Height -gt $innerH + 1.5) {
                [void]$results.Add(("{0} | {1} | '{2}' | tekst wyższy niż miejsce: {3:N0}px > {4:N0}px" -f $WindowName, $ownerLabel, $textShort, $ft.Height, $innerH))
            }
        } else {
            if ($tb.TextTrimming -eq [System.Windows.TextTrimming]::None -and $ft.WidthIncludingTrailingWhitespace -gt $innerW + 1.5) {
                [void]$results.Add(("{0} | {1} | '{2}' | ucięty w poziomie: tekst {3:N0}px, miejsce {4:N0}px" -f $WindowName, $ownerLabel, $textShort, $ft.WidthIncludingTrailingWhitespace, $innerW))
            }
            if ($ft.Height -gt $innerH + 1.5) {
                [void]$results.Add(("{0} | {1} | '{2}' | ucięty w pionie: tekst {3:N0}px, miejsce {4:N0}px" -f $WindowName, $ownerLabel, $textShort, $ft.Height, $innerH))
            }
        }

        # 2) Czy któryś z rodziców nie przycina TextBlocka (do najbliższego ScrollViewera).
        $tbRect = New-Object System.Windows.Rect 0, 0, $tb.ActualWidth, $tb.ActualHeight
        $a = [System.Windows.Media.VisualTreeHelper]::GetParent($tb)
        while ($null -ne $a -and $a -ne $root) {
            if ($a -is [System.Windows.Controls.ScrollContentPresenter]) { break }
            if ($a -is [System.Windows.FrameworkElement]) {
                $clip = [System.Windows.Controls.Primitives.LayoutInformation]::GetLayoutClip($a)
                $clipRect = $null
                if ($null -ne $clip) { $clipRect = $clip.Bounds }
                elseif ($a.ClipToBounds) { $clipRect = New-Object System.Windows.Rect 0, 0, $a.ActualWidth, $a.ActualHeight }
                if ($null -ne $clipRect -and -not $clipRect.IsEmpty) {
                    try {
                        $r = $tb.TransformToAncestor($a).TransformBounds($tbRect)
                        $cutX = [Math]::Max(0, $clipRect.Left - $r.Left) + [Math]::Max(0, $r.Right - $clipRect.Right)
                        $cutY = [Math]::Max(0, $clipRect.Top - $r.Top) + [Math]::Max(0, $r.Bottom - $clipRect.Bottom)
                        if ($cutX -gt 1.5 -or $cutY -gt 1.5) {
                            [void]$results.Add(("{0} | {1} | '{2}' | przycięty przez {3}: brakuje {4:N0}px w poziomie, {5:N0}px w pionie" -f $WindowName, $ownerLabel, $textShort, (Get-UiElementLabel $a), $cutX, $cutY))
                            break
                        }
                    } catch {}
                }
            }
            $a = [System.Windows.Media.VisualTreeHelper]::GetParent($a)
        }
    }
    return ,$results
}

# Tekstowa "mapa" okna: widoczne kontrolki z pozycją i rozmiarem względem okna.
function global:Write-UiLayoutMap {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $root = [System.Windows.Media.VisualTreeHelper]::GetChild($Win, 0)
    Write-Host ("MAP  == {0} ({1:N0} x {2:N0}) ==" -f $WindowName, $root.ActualWidth, $root.ActualHeight)
    $stack = New-Object System.Collections.Stack
    $stack.Push($root)
    while ($stack.Count -gt 0) {
        $v = $stack.Pop()
        if ($v -is [System.Windows.UIElement] -and -not $v.IsVisible) { continue }
        $isControl = $v -is [System.Windows.Controls.Button] -or $v -is [System.Windows.Controls.TextBox] -or $v -is [System.Windows.Controls.PasswordBox] -or
            $v -is [System.Windows.Controls.ComboBox] -or $v -is [System.Windows.Controls.CheckBox] -or $v -is [System.Windows.Controls.ListView] -or
            $v -is [System.Windows.Controls.ListBox] -or $v -is [System.Windows.Controls.ProgressBar]
        $isLooseText = $v -is [System.Windows.Controls.TextBlock] -and $null -eq $v.TemplatedParent -and -not [string]::IsNullOrWhiteSpace($v.Text)
        if ($isControl -or $isLooseText) {
            try {
                $r = $v.TransformToAncestor($root).TransformBounds((New-Object System.Windows.Rect 0, 0, $v.ActualWidth, $v.ActualHeight))
                $text = ''
                if ($v -is [System.Windows.Controls.TextBlock]) { $text = $v.Text }
                elseif ($v -is [System.Windows.Controls.ContentControl] -and $v.Content -is [string]) { $text = $v.Content }
                elseif ($v -is [System.Windows.Controls.TextBox]) { $text = $v.Text }
                elseif ($v -is [System.Windows.Controls.ComboBox]) { $text = [string]$v.Text }
                $text = ($text -replace '\s+', ' ').Trim()
                if ($text.Length -gt 50) { $text = $text.Substring(0, 50) + '…' }
                Write-Host ("MAP  {0,-34} x={1,4:N0} y={2,4:N0} w={3,4:N0} h={4,3:N0} fs={5} '{6}'" -f (Get-UiElementLabel $v), $r.X, $r.Y, $r.Width, $r.Height, $v.FontSize, $text)
            } catch {}
            if ($isControl -and $v -isnot [System.Windows.Controls.ListView] -and $v -isnot [System.Windows.Controls.ListBox]) { continue }
        }
        for ($i = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v) - 1; $i -ge 0; $i--) {
            $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i))
        }
    }
}

# ---------- Kontrole wyglądu: STALE / GLYPH / EMOJI (lista w uwagi-motyw.txt) ----------
# Ten sam wzorzec emoji co w teście Pester "Teksty okien bez emoji"; strzałki sortowania (U+25B2/25BC)
# celowo poza nim - to zwykłe kształty z Segoe UI.
$global:UiEmojiPattern = '[\u2300-\u23FF\u2600-\u27BF\u2B00-\u2BFF\u25B6\u25BE\u2139\uFE0F]|[\uD83C-\uDBFF][\uDC00-\uDFFF]'

# Widoczne elementy drzewa wizualnego (ukryte gałęzie pominięte). Zwracamy ,$lista - bez @() wokół
# listy generycznej (w PS 5.1 to błąd "Argument types do not match").
function global:Get-UiVisibleVisuals {
    param($Root)
    $list = New-Object System.Collections.Generic.List[object]
    $stack = New-Object System.Collections.Stack
    $stack.Push($Root)
    while ($stack.Count -gt 0) {
        $v = $stack.Pop()
        if ($v -is [System.Windows.UIElement] -and -not $v.IsVisible) { continue }
        $list.Add($v)
        for ($i = [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v) - 1; $i -ge 0; $i--) {
            $stack.Push([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i))
        }
    }
    return ,$list
}

# Opis elementu do raportu: Typ#nazwa; element bez własnej nazwy z XAML okna (np. część szablonu
# albo wiersza listy) dostaje jeszcze najbliższego nazwanego przodka ("Border w ListView#lvApps"),
# a TextBlock - swój tekst.
function global:Get-UiElementPath {
    param($Element)
    $label = Get-UiElementLabel $Element
    $isNamed = $Element -is [System.Windows.FrameworkElement] -and -not [string]::IsNullOrWhiteSpace($Element.Name) -and $null -eq $Element.TemplatedParent
    if (-not $isNamed) {
        $p = [System.Windows.Media.VisualTreeHelper]::GetParent($Element)
        while ($null -ne $p) {
            if ($p -is [System.Windows.FrameworkElement] -and -not [string]::IsNullOrWhiteSpace($p.Name) -and $null -eq $p.TemplatedParent) { $label += " w $(Get-UiElementLabel $p)"; break }
            $p = [System.Windows.Media.VisualTreeHelper]::GetParent($p)
        }
    }
    if ($Element -is [System.Windows.Controls.TextBlock] -and -not [string]::IsNullOrWhiteSpace($Element.Text)) { $label += " '$(Get-UiShortText $Element.Text)'" }
    return $label
}

function global:Get-UiShortText {
    param([string]$Text, [int]$Max = 50)
    $t = ($Text -replace '\s+', ' ').Trim()
    if ($t.Length -gt $Max) { $t = $t.Substring(0, $Max) + '…' }
    return $t
}

# Kod znaku do raportu (U+XXXX) - konsola CI w PS 5.1 nie zawsze wypisze sam znak.
function global:Get-UiCodePoint {
    param([string]$Text)
    if ([string]::IsNullOrEmpty($Text)) { return '' }
    if ($Text.Length -ge 2 -and [char]::IsSurrogatePair($Text[0], $Text[1])) { return ('U+{0:X4}' -f [char]::ConvertToUtf32($Text[0], $Text[1])) }
    return ('U+{0:X4}' -f [int]$Text[0])
}

# Uwaga trafia na listę (i od razu do konsoli) tylko raz na okno: ten sam element widoczny w każdej
# zakładce Ustawień ("ustawienia-01" ... "-09") dałby inaczej dziewięć identycznych wpisów.
$global:UiFindingKeys = @{}
function global:Add-UiFinding {
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return }
    $key = $Text -replace '^(\w+ [^|]*?)-\d\d \|', '$1 |'
    if ($global:UiFindingKeys.ContainsKey($key)) { return }
    $global:UiFindingKeys[$key] = $true
    [void]$global:UiFindings.Add($Text)
    Write-Host $Text
}

# GLYPH: znak ikony (obszar prywatny Unicode, U+E000-U+F8FF) w TextBlocku, który nie ma czcionki ikon -
# rysuje się jako pusty prostokąt.
function global:Find-UiGlyphFont {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $results = New-Object System.Collections.Generic.List[string]
    foreach ($v in (Get-UiVisibleVisuals -Root $Win)) {
        if ($v -isnot [System.Windows.Controls.TextBlock] -or [string]::IsNullOrEmpty($v.Text)) { continue }
        $m = [regex]::Match($v.Text, '[\uE000-\uF8FF]')
        if ($m.Success -and $v.FontFamily.Source -notmatch 'MDL2') {
            $results.Add(("GLYPH {0} | {1} | {2} | czcionka '{3}'" -f $WindowName, (Get-UiElementPath $v), (Get-UiCodePoint $m.Value), $v.FontFamily.Source))
        }
    }
    return ,$results
}

# EMOJI: emoji w widocznym tekście okna (WPF rysuje je czcionką zastępczą: inny rozmiar i kolor,
# w motywie jasnym bywają nieczytelne). Zamiast nich ikony z czcionki ikon albo zwykły tekst.
function global:Find-UiEmoji {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $results = New-Object System.Collections.Generic.List[string]
    foreach ($v in (Get-UiVisibleVisuals -Root $Win)) {
        if ($v -isnot [System.Windows.Controls.TextBlock] -or [string]::IsNullOrEmpty($v.Text)) { continue }
        $m = [regex]::Match($v.Text, $global:UiEmojiPattern)
        if ($m.Success) {
            $results.Add(("EMOJI {0} | {1} | {2}" -f $WindowName, (Get-UiElementPath $v), (Get-UiCodePoint $m.Value)))
        }
    }
    return ,$results
}

# CONTRAST: widoczny tekst, którego kolor ma za mały kontrast z faktycznym tłem (pierwsze nieprzezroczyste
# tło w górę drzewa wizualnego, z mieszaniem półprzezroczystych warstw). Łapie np. czarny tekst
# odziedziczony z domyślnego stylu Windows na ciemnej karcie. Pomijamy tekst celowo przygaszony
# (kolor ThemeFaint: podpowiedzi w polach, wersja) i tekst w wyłączonych / półprzezroczystych elementach.
$global:UiMinContrast = 3.0
function global:Get-UiRelativeLuminance {
    param([System.Windows.Media.Color]$Color)
    $ch = foreach ($c in @($Color.R, $Color.G, $Color.B)) {
        $v = $c / 255.0
        if ($v -le 0.03928) { $v / 12.92 } else { [Math]::Pow(($v + 0.055) / 1.055, 2.4) }
    }
    return 0.2126 * $ch[0] + 0.7152 * $ch[1] + 0.0722 * $ch[2]
}
function global:Get-UiContrastRatio {
    param([System.Windows.Media.Color]$A, [System.Windows.Media.Color]$B)
    $la = Get-UiRelativeLuminance $A; $lb = Get-UiRelativeLuminance $B
    $hi = [Math]::Max($la, $lb); $lo = [Math]::Min($la, $lb)
    return ($hi + 0.05) / ($lo + 0.05)
}
# Tło pod elementem: pędzle Background (Panel, Border, Control - TextBlock też) w górę drzewa aż do
# pierwszego nieprzezroczystego; półprzezroczyste nakładamy na nie od dołu. $null = tło z obrazu.
function global:Get-UiEffectiveBackground {
    param($Element, [System.Windows.Window]$Win)
    $layers = New-Object System.Collections.Generic.List[object]
    $v = $Element
    while ($null -ne $v) {
        $brush = $null
        if ($v -is [System.Windows.Controls.Panel]) { $brush = $v.Background }
        elseif ($v -is [System.Windows.Controls.Border]) { $brush = $v.Background }
        elseif ($v -is [System.Windows.Controls.TextBlock]) { $brush = $v.Background }
        elseif ($v -is [System.Windows.Window]) { $brush = $v.Background }
        $c = $null
        if ($brush -is [System.Windows.Media.SolidColorBrush]) { $c = $brush.Color }
        elseif ($brush -is [System.Windows.Media.GradientBrush] -and $brush.GradientStops.Count -gt 0) {
            # Gradient (np. kafelek logo w nagłówku): średnia kolorów punktów - znak stoi na środku kafelka.
            $r = 0; $g = 0; $bl = 0; $al = 0
            foreach ($stop in $brush.GradientStops) { $r += $stop.Color.R; $g += $stop.Color.G; $bl += $stop.Color.B; $al += $stop.Color.A }
            $n = $brush.GradientStops.Count
            $c = [System.Windows.Media.Color]::FromArgb([byte][Math]::Round($al / $n), [byte][Math]::Round($r / $n), [byte][Math]::Round($g / $n), [byte][Math]::Round($bl / $n))
        }
        elseif ($null -ne $brush) {
            # Obraz / VisualBrush: tła nie da się policzyć z kolorów, więc tekstu nie oceniamy.
            return $null
        }
        if ($null -ne $c) {
            $a = $c.A * $brush.Opacity
            if ($a -ge 255) { $layers.Add($c); break }
            if ($a -gt 0) { $layers.Add([System.Windows.Media.Color]::FromArgb([byte][Math]::Round($a), $c.R, $c.G, $c.B)) }
        }
        $v = [System.Windows.Media.VisualTreeHelper]::GetParent($v)
    }
    $result = [System.Windows.Media.Colors]::White
    if ($layers.Count -gt 0 -and $layers[$layers.Count - 1].A -ge 255) { $result = $layers[$layers.Count - 1] }
    elseif ($Win.Background -is [System.Windows.Media.SolidColorBrush]) { $result = $Win.Background.Color }
    for ($i = $layers.Count - 1; $i -ge 0; $i--) {
        $l = $layers[$i]
        if ($l.A -ge 255) { $result = $l; continue }
        $f = $l.A / 255.0
        $result = [System.Windows.Media.Color]::FromRgb([byte][Math]::Round($l.R * $f + $result.R * (1 - $f)), [byte][Math]::Round($l.G * $f + $result.G * (1 - $f)), [byte][Math]::Round($l.B * $f + $result.B * (1 - $f)))
    }
    return $result
}
function global:Find-UiContrast {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $results = New-Object System.Collections.Generic.List[string]
    $faint = ([string](Get-ThemePalette)['ThemeFaint']).ToUpperInvariant()
    foreach ($v in (Get-UiVisibleVisuals -Root $Win)) {
        if ($v -isnot [System.Windows.Controls.TextBlock] -or [string]::IsNullOrWhiteSpace($v.Text)) { continue }
        if ($v.Foreground -isnot [System.Windows.Media.SolidColorBrush]) { continue }
        $fg = $v.Foreground.Color
        $fgHex = ('#{0:X2}{1:X2}{2:X2}' -f $fg.R, $fg.G, $fg.B)
        if ($fgHex -eq $faint) { continue }
        $dimmed = $false
        $p = $v
        while ($null -ne $p) {
            if (($p -is [System.Windows.UIElement]) -and ($p.Opacity -lt 0.99 -or -not $p.IsEnabled)) { $dimmed = $true; break }
            $p = [System.Windows.Media.VisualTreeHelper]::GetParent($p)
        }
        if ($dimmed) { continue }
        $bg = Get-UiEffectiveBackground -Element $v -Win $Win
        if ($null -eq $bg) { continue }
        $ratio = Get-UiContrastRatio $fg $bg
        if ($ratio -lt $global:UiMinContrast) {
            $results.Add(("CONTRAST {0} | {1} | tekst {2} na tle #{3:X2}{4:X2}{5:X2} = {6:N2}:1 (min {7})" -f $WindowName, (Get-UiElementPath $v), $fgHex, $bg.R, $bg.G, $bg.B, $ratio, $global:UiMinContrast))
        }
    }
    return ,$results
}

# STALE: po przełączeniu motywu w locie żaden widoczny element nie może mieć koloru, który jest tylko
# w poprzedniej palecie (stare wartości minus nowe). Taki kolor to znak {StaticResource Theme...},
# koloru HEX w danych wiersza albo pędzla przypisanego w kodzie (New-ThemeBrush) zamiast
# SetResourceReference. Palety jako IDictionary: [ordered] nie wiąże się z parametrem [hashtable].
# Foreground sprawdzamy tylko tam, gdzie jest ustawiony (lokalnie, stylem, szablonem) - wartość
# dziedziczona trafia do raportu raz, przy elemencie, który ją ustawia.
function global:Find-UiStaleColors {
    param(
        [System.Windows.Window]$Win,
        [string]$WindowName,
        [System.Collections.IDictionary]$OldPalette,
        [System.Collections.IDictionary]$NewPalette
    )
    $results = New-Object System.Collections.Generic.List[string]
    $newHex = @{}
    foreach ($k in $NewPalette.Keys) { $newHex[([string]$NewPalette[$k]).ToUpperInvariant()] = $true }
    $staleHex = @{}
    foreach ($k in $OldPalette.Keys) {
        $hex = ([string]$OldPalette[$k]).ToUpperInvariant()
        if ($newHex.ContainsKey($hex)) { continue }
        if ($staleHex.ContainsKey($hex)) { $staleHex[$hex] = $staleHex[$hex] + ', ' + $k } else { $staleHex[$hex] = [string]$k }
    }
    if ($staleHex.Count -eq 0) { return ,$results }

    $fgDp = [System.Windows.Documents.TextElement]::ForegroundProperty
    $counts = @{}
    $order = New-Object System.Collections.Generic.List[string]
    foreach ($v in (Get-UiVisibleVisuals -Root $Win)) {
        $pairs = @()
        if ($v -is [System.Windows.Controls.Control] -or $v -is [System.Windows.Controls.Border]) { $pairs = @(@('Background', $v.Background), @('BorderBrush', $v.BorderBrush)) }
        elseif ($v -is [System.Windows.Controls.Panel] -or $v -is [System.Windows.Controls.TextBlock]) { $pairs = @(, @('Background', $v.Background)) }
        elseif ($v -is [System.Windows.Shapes.Shape]) { $pairs = @(@('Fill', $v.Fill), @('Stroke', $v.Stroke)) }
        if ($v -is [System.Windows.FrameworkElement]) {
            $src = [System.Windows.DependencyPropertyHelper]::GetValueSource($v, $fgDp).BaseValueSource
            if ($src -ne [System.Windows.BaseValueSource]::Inherited -and $src -ne [System.Windows.BaseValueSource]::Default) { $pairs += , @('Foreground', $v.GetValue($fgDp)) }
        }
        foreach ($pair in $pairs) {
            $brush = $pair[1]
            if ($brush -isnot [System.Windows.Media.SolidColorBrush]) { continue }
            $col = $brush.Color
            if ($col.A -eq 0) { continue }
            $hex = '#{0:X2}{1:X2}{2:X2}' -f $col.R, $col.G, $col.B
            if (-not $staleHex.ContainsKey($hex)) { continue }
            $line = "STALE {0} | {1} | {2} | {3} ({4})" -f $WindowName, (Get-UiElementPath $v), $pair[0], $hex, $staleHex[$hex]
            if ($counts.ContainsKey($line)) { $counts[$line]++ } else { $counts[$line] = 1; $order.Add($line) }
        }
    }
    foreach ($line in $order) {
        if ($counts[$line] -gt 1) { $results.Add("$line | x$($counts[$line])") } else { $results.Add($line) }
    }
    return ,$results
}

# Przełącza motyw na otwartym oknie tak jak przycisk motywu (Set-AppTheme; okno bez właściciela,
# np. logowanie albo okno dialogowe po zamknięciu okna głównego, dostaje Update-WindowTheme wprost),
# szuka kolorów z poprzedniego motywu i zawsze wraca do motywu wyjściowego. Zrzut PNG jest już zrobiony.
function global:Test-UiThemeSwitch {
    param([System.Windows.Window]$Win, [string]$WindowName)
    $wasDark = [bool]$script:isDarkTheme
    $oldPalette = Get-ThemePalette
    $found = $null
    try {
        $script:isDarkTheme = -not $wasDark
        Set-AppTheme
        Update-WindowTheme $Win
        $Win.UpdateLayout()
        $found = Find-UiStaleColors -Win $Win -WindowName $WindowName -OldPalette $oldPalette -NewPalette (Get-ThemePalette)
    } catch {
        Write-Host "ERR  ${WindowName} (przełączenie motywu): $($_.Exception.Message)"
        [void]$global:UiErrors.Add("${WindowName} (przełączenie motywu): $($_.Exception.Message)")
    } finally {
        $script:isDarkTheme = $wasDark
        try { Set-AppTheme } catch {}
        Update-WindowTheme $Win
        $Win.UpdateLayout()
    }
    if ($null -ne $found) { foreach ($issue in $found) { Add-UiFinding $issue } }
}

# Jeden "widok" okna: lint ucinania, mapa układu, PNG i (gdy jest przewijanie) PNG całej zawartości.
# -Lint: dodatkowo GLYPH / EMOJI i przełączenie motywu w locie (STALE) - dla zwykłych zrzutów i stanów
# dodatkowych; zrzuty w innym rozmiarze sprawdzają tylko ucinanie (reszta się od rozmiaru nie zmienia).
function global:Save-UiView {
    param([System.Windows.Window]$Win, [string]$Name, [switch]$Lint)
    $root = [System.Windows.Media.VisualTreeHelper]::GetChild($Win, 0)
    $root.UpdateLayout()
    foreach ($issue in (Find-UiClipping -Win $Win -WindowName $Name)) {
        [void]$global:UiClipping.Add($issue)
        Write-Host "CLIP $issue"
    }
    if ($global:UiDumpLayout) { Write-UiLayoutMap -Win $Win -WindowName $Name }
    Save-UiElementPng -Element $root -Background $null -Name $Name
    # Jeśli okno ma przewijaną zawartość dłuższą niż widok - dodatkowo cała zawartość.
    $queue = New-Object System.Collections.Queue
    $queue.Enqueue($root)
    while ($queue.Count -gt 0) {
        $v = $queue.Dequeue()
        if ($v -is [System.Windows.UIElement] -and -not $v.IsVisible) { continue }
        if ($v -is [System.Windows.Controls.ScrollViewer] -and $v.ExtentHeight -gt ($v.ViewportHeight + 1) -and $v.ViewportHeight -gt 0 -and $v.TemplatedParent -isnot [System.Windows.Controls.Primitives.TextBoxBase] -and $v.TemplatedParent -isnot [System.Windows.Controls.ItemsControl]) {
            # Zawartość dłuższa niż widoczny obszar = trzeba przewijać (np. lista zadań w oknie głównym).
            Write-Host ("SCROLL {0} | {1} | zawartość {2:N0}px w widoku {3:N0}px" -f $Name, (Get-UiElementLabel $v.Content), $v.ExtentHeight, $v.ViewportHeight)
        }
        if ($v -is [System.Windows.Controls.ScrollViewer] -and $v.Content -is [System.Windows.FrameworkElement] -and $v.ExtentHeight -gt ($v.ViewportHeight + 40) -and $v.ViewportHeight -gt 100) {
            Save-UiElementPng -Element $v.Content -Background $Win.Background -Name "$Name-cala-zawartosc"
            continue
        }
        for ($i = 0; $i -lt [System.Windows.Media.VisualTreeHelper]::GetChildrenCount($v); $i++) {
            $queue.Enqueue([System.Windows.Media.VisualTreeHelper]::GetChild($v, $i))
        }
    }
    if ($Lint) {
        foreach ($issue in (Find-UiGlyphFont -Win $Win -WindowName $Name)) { Add-UiFinding $issue }
        foreach ($issue in (Find-UiEmoji -Win $Win -WindowName $Name)) { Add-UiFinding $issue }
        foreach ($issue in (Find-UiContrast -Win $Win -WindowName $Name)) { Add-UiFinding $issue }
        Test-UiThemeSwitch -Win $Win -WindowName $Name
    }
}

# Okno z zakładkami: każda zakładka osobno (w drzewie wizualnym jest tylko wybrana).
function global:Save-UiTabbedView {
    param([System.Windows.Window]$Win, [string]$Name, [switch]$Lint)
    $tabs = $null
    foreach ($tabName in 'tabSettings', 'tabSysInfo') { if ($null -eq $tabs) { $tabs = $Win.FindName($tabName) } }
    if ($null -eq $tabs) { Save-UiView -Win $Win -Name $Name -Lint:$Lint; return }
    for ($i = 0; $i -lt $tabs.Items.Count; $i++) {
        $tabs.SelectedIndex = $i
        $viewName = "{0}-{1:D2}" -f $Name, ($i + 1)
        Write-Host ("TAB  {0} = {1}" -f $viewName, $tabs.Items[$i].Header)
        Save-UiView -Win $Win -Name $viewName -Lint:$Lint
    }
    $tabs.SelectedIndex = 0
}

# Zrzut w podanym rozmiarze okna (DIP). MinWidth/MinHeight zerujemy, bo mały ekran jest mniejszy niż
# minimalny rozmiar okna - narzędzie samo przycina je wtedy do obszaru roboczego (New-ThemedWindow).
function global:Save-UiSizedView {
    param([System.Windows.Window]$Win, [string]$Name, [double]$Width, [double]$Height)
    $Win.MinWidth = 0
    $Win.MinHeight = 0
    $Win.Width = $Width
    $Win.Height = $Height
    $Win.UpdateLayout()
    Save-UiTabbedView -Win $Win -Name $Name
}

function global:Save-UiWindowShot {
    param([System.Windows.Window]$Win)
    try {
        $name = $global:UiShotName
        Save-UiTabbedView -Win $Win -Name $name -Lint
        # Dodatkowe stany tego samego okna (np. okno główne w trakcie wdrożenia), w zwykłym rozmiarze.
        if ($global:UiExtraViews.ContainsKey($name)) {
            foreach ($view in $global:UiExtraViews[$name]) {
                try {
                    & $view.Apply | Out-Null
                    $Win.UpdateLayout()
                    Save-UiView -Win $Win -Name $view.Name -Lint
                } catch {
                    Write-Host "ERR  $($view.Name): $($_.Exception.Message)"
                    [void]$global:UiErrors.Add("$($view.Name): $($_.Exception.Message)")
                } finally {
                    try { & $view.Revert | Out-Null } catch {
                        Write-Host "ERR  $($view.Name) (powrót do stanu wyjściowego): $($_.Exception.Message)"
                        [void]$global:UiErrors.Add("$($view.Name) (powrót do stanu wyjściowego): $($_.Exception.Message)")
                    }
                    $Win.UpdateLayout()
                }
            }
        }
        if ($name -in @('okno-glowne', 'ustawienia')) {
            # Minimalny rozmiar okna - tam najłatwiej coś uciąć.
            if ($Win.MinWidth -gt 0 -and $Win.MinHeight -gt 0) {
                $Win.Width = $Win.MinWidth
                $Win.Height = $Win.MinHeight
                Save-UiTabbedView -Win $Win -Name "$name-min"
            }
            # Ekran 1366x768 przy skalowaniu 150% bez paska zadań: 911 x 472 DIP. Tu nie może być żadnego
            # CLIP (ucięty tekst albo przycisk); SCROLL (przewijana lista zadań) jest oczekiwany.
            Save-UiSizedView -Win $Win -Name "$name-1366x768-150proc" -Width 911 -Height 472
        }
    } catch {
        Write-Host "ERR  $($global:UiShotName): $($_.Exception.Message)"
        [void]$global:UiErrors.Add("$($global:UiShotName): $($_.Exception.Message)")
    }
}

# Każde okno po wczytaniu (Loaded) czeka, aż dispatcher skończy pracę (dane, układ), po czym jest
# renderowane do PNG i zamykane.
$global:UiLoadedHandler = [System.Windows.RoutedEventHandler]{
    param($sender, $e)
    [void]$global:UiPending.Add($sender)
    [void]$sender.Dispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::ApplicationIdle, [Action]{
        while ($global:UiPending.Count -gt 0) {
            $w = $global:UiPending[0]
            $global:UiPending.RemoveAt(0)
            Save-UiWindowShot -Win $w
            try { $w.Close() } catch {}
        }
    })
}
[System.Windows.EventManager]::RegisterClassHandler([System.Windows.Window], [System.Windows.FrameworkElement]::LoadedEvent, $global:UiLoadedHandler)

# Przetwarza zdarzenia dispatchera, aż warunek będzie spełniony (dla okien niemodalnych).
function global:Wait-UiCondition {
    param([scriptblock]$Until, [int]$TimeoutMs = 20000)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while (-not (& $Until) -and $sw.ElapsedMilliseconds -lt $TimeoutMs) {
        $frame = New-Object System.Windows.Threading.DispatcherFrame
        [void][System.Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([System.Windows.Threading.DispatcherPriority]::SystemIdle, [Action]{ $frame.Continue = $false })
        [System.Windows.Threading.Dispatcher]::PushFrame($frame)
        Start-Sleep -Milliseconds 20
    }
}

function Invoke-UiShot {
    param([string]$Name, [scriptblock]$Action)
    $global:UiShotName = $Name
    try { & $Action | Out-Null } catch {
        Write-Host "ERR  ${Name}: $($_.Exception.Message)"
        [void]$global:UiErrors.Add("${Name}: $($_.Exception.Message)")
    }
}

# --- Dodatkowe stany okna głównego (po zwykłym zrzucie, na tym samym otwartym oknie) ---
# Kontrolki i funkcje są z narzędzia (wczytanego wyżej), tak jak w Start-Deployment i obsłudze Pauzy.
# Revert zawsze wraca do stanu "gotowy" - także gdy Apply przerwał się w połowie.
$global:UiApplyDeployState = {
    $btnStart.IsEnabled = $false
    $btnStart.Content = 'Trwa konfiguracja...'
    foreach ($deployButton in @($btnPause, $btnCancelDeploy)) {
        $deployButton.IsEnabled = $true
        $deployButton.Visibility = [System.Windows.Visibility]::Visible
    }
    # Jak w Step-DeploymentProgress: +1/-1, żeby pasek nie animował się od zera.
    $progressBar.Value = 43
    $progressBar.Value = 42
    Set-ProgressText 'Instalacja: Google Chrome (3 z 7)...'
    $txtStopwatch.Text = '00:04:12'
    $txtStopwatch.Visibility = [System.Windows.Visibility]::Visible
    $progressBarDownload.Value = 65
    $progressBarDownload.Visibility = [System.Windows.Visibility]::Visible
    Set-DeployStatusDot Busy
}
$global:UiResetMainWindow = {
    foreach ($deployButton in @($btnPause, $btnCancelDeploy)) {
        $deployButton.IsEnabled = $false
        $deployButton.Visibility = [System.Windows.Visibility]::Collapsed
    }
    Set-ButtonContent $btnPause 'Pauza' 'E769'
    $btnPause.Style = $Window.FindResource('WarningButton')
    $btnStart.IsEnabled = $true
    & $syncStartLabel
    $progressBar.Value = 0
    Set-ProgressText 'Gotowy do rozpoczęcia'
    $txtStopwatch.Text = '00:00:00'
    $txtStopwatch.Visibility = [System.Windows.Visibility]::Collapsed
    $progressBarDownload.Value = 0
    $progressBarDownload.Visibility = [System.Windows.Visibility]::Hidden
    Set-DeployStatusDot Idle
    $mainToastHost = $Window.FindName('toastHost')
    if ($null -ne $mainToastHost) { $mainToastHost.Children.Clear() }
}
$global:UiExtraViews['okno-glowne'] = @(
    @{
        Name   = 'okno-glowne-wdrozenie'
        Apply  = { & $global:UiApplyDeployState }
        Revert = { & $global:UiResetMainWindow }
    },
    @{
        # Pauza: przycisk "Wznów" w miękkim zielonym stylu, bursztynowa kropka (jak w obsłudze btnPause).
        Name   = 'okno-glowne-pauza'
        Apply  = {
            & $global:UiApplyDeployState
            Set-ButtonContent $btnPause 'Wznów' 'E768'
            $btnPause.Style = $Window.FindResource('SuccessButton')
            Set-DeployStatusDot Paused
        }
        Revert = { & $global:UiResetMainWindow }
    },
    @{
        # Powiadomienia: -Seconds 30, żeby nie zniknęły przed zrzutem (w trybie testowym bez animacji).
        Name   = 'okno-glowne-powiadomienia'
        Apply  = {
            Show-Toast -Text 'Zapisano konfigurację.' -Kind Ok -Seconds 30 -TargetWindow $Window
            Show-Toast -Text 'Nie udało się sprawdzić aktualizacji: ścieżka \\serwer\deploy\std_version.json jest niedostępna.' -Kind Crit -Seconds 30 -TargetWindow $Window
        }
        Revert = { & $global:UiResetMainWindow }
    }
)

# --- Motyw ---
$script:isDarkTheme = ($Theme -eq 'Dark')
Set-AppTheme
foreach ($w in @($authWindow, $welcomeWindow)) { if ($null -ne $w) { Update-WindowTheme $w } }

# --- Przykładowe wpisy logu (różne rodzaje), żeby przeglądarka logów nie była pusta ---
Write-Log "Rozpoczęto walidację konfiguracji i plików przed wdrożeniem..."
Write-Log "Mało wolnego miejsca na dysku C: (12.4 GB). Instalacja dużych programów może zakończyć się błędem."
Write-Log "[DRY-RUN] Zainstalowano by Google Chrome przez Winget: winget install --id `"Google.Chrome`" -e --silent"
Write-Log "Instalator 7-Zip zakończył się błędem (kod wyjścia: 1603)." -IsError
Write-Log "7-Zip zainstalowany." -Context "Automat"
Write-Log "Zastosowano profil wdrożenia: Standard (Wybrano programów: 3)" -Context "Użytkownik"

# --- Okna ---
Get-AppSelection
Load-Profiles
$cfg = Get-Config

Invoke-UiShot 'logowanie' { $authWindow.ShowDialog() }
Invoke-UiShot 'powitanie' { $welcomeWindow.ShowDialog() }
Invoke-UiShot 'okno-glowne' { $Window.ShowDialog() }

Invoke-UiShot 'komunikat-pytanie' { Show-ThemedMessageBox -Message "Czy na pewno chcesz rozpocząć konfigurację?" -Title "Potwierdzenie" -Button "YesNo" -Image "Question" }
Invoke-UiShot 'komunikat-blad' { Show-ThemedMessageBox -Message ("Wykryto bledy walidacji:`r`n- Brak sciezki sieciowej (InstallSourcePaths.network).`r`n- Nie znaleziono pliku dla 'Google Chrome': \\serwer\udzial\ChromeSetup.exe`r`n- Brak DomainJoin.DomainName w config.json.") -Title "Walidacja" -Button "OK" -Image "Error" }
Invoke-UiShot 'komunikat-ostrzezenie' { Show-ThemedMessageBox -Message "Czy na pewno chcesz przerwać wdrożenie?" -Title "Przerwij" -Button "YesNoCancel" -Image "Warning" }
Invoke-UiShot 'komunikat-sukces' { Show-ThemedMessageBox -Message "Wyeksportowano raport do C:\Users\Technik\Desktop\Raport_logow_DESKTOP-AB12CD3.html" -Title "Sukces" -Button "OK" -Image "Success" }
Invoke-UiShot 'wpisz-tekst' { Show-InputDialog -Title "Zmiana nazwy komputera" -Message "Podaj nową nazwę komputera (maks. 15 znaków). Obecna nazwa: DESKTOP-AB12CD3" -DefaultText "PC-5CG1234XYZ" }
Invoke-UiShot 'wpisz-haslo' { Show-InputDialog -Title "Konto lokalnego administratora" -Message "Podaj hasło dla nowego konta 'LokalnyIT' (co najmniej 8 znaków):" -Password }
Invoke-UiShot 'pin' { Show-PinPrompt }
Invoke-UiShot 'wybor-aplikacji' { Show-AppSelectionWindow }
Invoke-UiShot 'informacja' { Show-CustomInfoDialog -Title "Zakończono" -Message "Przetwarzanie deinstalacji zakończone.`n`nPoprawnie odinstalowane aplikacje zostały automatycznie usunięte z listy.`nJeśli instalator zwrócił błąd, aplikacja pozostała na liście oznaczona jako «Błąd»." -ShowCopy -HtmlData "<html></html>" }
Invoke-UiShot 'potwierdzenie-deinstalacji' { Show-UninstallConfirmDialog -AppCount 3 -AppListText "- 7-Zip 23.01 (x64)`n- Google Chrome`n- Microsoft OneDrive" }
Invoke-UiShot 'logi' { Show-LogWindow; Wait-UiCondition -Until { $null -eq $script:ActiveLogWindow -or -not $script:ActiveLogWindow.IsVisible } }
Invoke-UiShot 'szczegoly-wpisu-logu' { Show-LogEntryDetail -Entry ([PSCustomObject]@{ RodzajText = 'Błąd'; RodzajKey = 'Error'; DataText = '06.10.2026 09:12:44'; Kontekst = 'Automat'; Informacja = "Instalator 7-Zip zakończył się błędem (kod wyjścia: 1603).`r`nSzczegóły: Fatal error during installation."; RowColorHex = (Get-ThemeColor 'ThemeDanger') }) }
Invoke-UiShot 'informacje-o-systemie' { Show-SystemInfoWindow }
Invoke-UiShot 'deinstalator' { Show-SoftwareUninstaller }
Invoke-UiShot 'ustawienia' { Show-ConfigEditor }
Invoke-UiShot 'edycja-programu' { Show-ProgramEditDialog -IsNew $false -ProgramName 'Google Chrome' -ProgramData ([PSCustomObject]@{ Enabled = $true; FileName = 'Google.Chrome'; SilentArgs = ''; DownloadUrl = 'https://dl.google.com/chrome/install/ChromeStandaloneSetup64.exe' }) }
Invoke-UiShot 'edycja-rejestru' { Show-RegistryEditDialog -IsNew $false -RegData ([PSCustomObject]@{ Path = 'HKLM:\SOFTWARE\MojaFirma'; Name = 'WdrozenieZakonczone'; Value = '1'; PropertyType = 'DWord' }) }
Invoke-UiShot 'edycja-skryptu' { Show-ScriptEditDialog -IsNew $true -ScriptPath 'SkryptDrukarki.ps1' }
Invoke-UiShot 'edycja-profilu' { Show-ProfileEditDialog -IsNew $false -ProfileName 'Standard' -ProfileApps @('7-Zip', 'Google Chrome') -config $cfg }

# Ikona narzędzia (New-AppIconBytes): każdy rozmiar z pliku .ico w skali 1:1 na tle motywu - wersja ze
# znakiem z czcionki i zapasowa rysowana kreską - oraz sam plik ikona.ico do obejrzenia w Eksploratorze.
try {
    [System.IO.File]::WriteAllBytes((Join-Path $global:UiShotDir 'ikona.ico'), $script:AppIconBytes)
    $textBrush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString((Get-ThemeColor 'ThemeText')))
    $iconPanel = New-Object System.Windows.Controls.StackPanel
    $iconPanel.Margin = '12'
    $variants = @(
        @{ Label = 'Ikona narzędzia (znak z czcionki Segoe Fluent Icons / Segoe MDL2 Assets)'; Bytes = $script:AppIconBytes },
        @{ Label = 'Zapas bez czcionki z ikonami (paczka rysowana kreską)'; Bytes = (New-AppIconBytes -DrawnMark) }
    )
    foreach ($variant in $variants) {
        $caption = New-Object System.Windows.Controls.TextBlock
        $caption.Text = $variant.Label
        $caption.Foreground = $textBrush
        $caption.Margin = '0,8,0,6'
        [void]$iconPanel.Children.Add($caption)
        $row = New-Object System.Windows.Controls.StackPanel
        $row.Orientation = 'Horizontal'
        $decoder = New-Object System.Windows.Media.Imaging.IconBitmapDecoder -ArgumentList (New-Object System.IO.MemoryStream (, $variant.Bytes)), ([System.Windows.Media.Imaging.BitmapCreateOptions]::None), ([System.Windows.Media.Imaging.BitmapCacheOption]::OnLoad)
        foreach ($frame in $decoder.Frames) {
            $img = New-Object System.Windows.Controls.Image
            $img.Source = $frame
            $img.Width = $frame.PixelWidth
            $img.Height = $frame.PixelHeight
            $img.Stretch = 'Fill'
            $img.VerticalAlignment = 'Bottom'
            $img.Margin = '0,0,12,0'
            [System.Windows.Media.RenderOptions]::SetBitmapScalingMode($img, [System.Windows.Media.BitmapScalingMode]::NearestNeighbor)
            [void]$row.Children.Add($img)
        }
        [void]$iconPanel.Children.Add($row)
    }
    $iconPanel.Measure((New-Object System.Windows.Size ([double]::PositiveInfinity), ([double]::PositiveInfinity)))
    $iconPanel.Arrange((New-Object System.Windows.Rect $iconPanel.DesiredSize))
    $bgBrush = New-Object System.Windows.Media.SolidColorBrush ([System.Windows.Media.ColorConverter]::ConvertFromString((Get-ThemeColor 'ThemeBackground')))
    Save-UiElementPng -Element $iconPanel -Background $bgBrush -Name 'ikona'
} catch {
    Write-Host "ERR  ikona: $($_.Exception.Message)"
    [void]$global:UiErrors.Add("ikona: $($_.Exception.Message)")
}

Write-Host "Gotowe: $global:UiShotIndex plików w $global:UiShotDir"

# --- Raporty ---
$report = Join-Path $global:UiShotDir 'uciety-tekst.txt'
if ($global:UiClipping.Count -gt 0) {
    $global:UiClipping | Set-Content -Path $report -Encoding UTF8
    Write-Host "UCIĘTY TEKST: $($global:UiClipping.Count) miejsc (lista w $report)"
} else {
    'Brak uciętego tekstu.' | Set-Content -Path $report -Encoding UTF8
    Write-Host "Brak uciętego tekstu."
}
$findingsReport = Join-Path $global:UiShotDir 'uwagi-motyw.txt'
if ($global:UiFindings.Count -gt 0) {
    $global:UiFindings | Set-Content -Path $findingsReport -Encoding UTF8
} else {
    'Brak uwag (STALE / GLYPH / EMOJI / CONTRAST).' | Set-Content -Path $findingsReport -Encoding UTF8
}

# --- Podsumowanie na końcu logu: w CI linie CLIP/STALE giną wśród linii MAP (-DumpLayout) ---
$findingCount = @{ STALE = 0; GLYPH = 0; EMOJI = 0; CONTRAST = 0 }
foreach ($finding in $global:UiFindings) {
    $kind = ([string]$finding).Split(' ')[0]
    if ($findingCount.ContainsKey($kind)) { $findingCount[$kind]++ }
}
$summary = New-Object System.Collections.Generic.List[string]
$summary.Add("==== Podgląd UI, motyw ${Theme}: uwagi ====")
$summary.Add("CLIP  ucięty tekst                  : $($global:UiClipping.Count)")
$summary.Add("STALE kolor z poprzedniego motywu   : $($findingCount.STALE)")
$summary.Add("GLYPH ikona w zwykłej czcionce      : $($findingCount.GLYPH)")
$summary.Add("EMOJI emoji w tekście okna          : $($findingCount.EMOJI)")
$summary.Add("CONTRAST za słaby kontrast tekstu   : $($findingCount.CONTRAST)")
$summary.Add("ERR   błąd przy renderowaniu okna   : $($global:UiErrors.Count)")
foreach ($clipLine in $global:UiClipping) { $summary.Add("CLIP $clipLine") }
foreach ($finding in $global:UiFindings) { $summary.Add([string]$finding) }
foreach ($errLine in $global:UiErrors) { $summary.Add("ERR  $errLine") }
$summary.Add("==== koniec uwag (pliki: $report, $findingsReport) ====")
foreach ($line in $summary) { Write-Host $line }

# GitHub Actions: liczniki jako adnotacja przebiegu i lista w podsumowaniu zadania (Job summary).
if ($env:GITHUB_ACTIONS -eq 'true') {
    $totals = "CLIP $($global:UiClipping.Count), STALE $($findingCount.STALE), GLYPH $($findingCount.GLYPH), EMOJI $($findingCount.EMOJI), CONTRAST $($findingCount.CONTRAST), ERR $($global:UiErrors.Count)"
    if (($global:UiClipping.Count + $global:UiFindings.Count + $global:UiErrors.Count) -gt 0) {
        Write-Host "::warning title=Podgląd UI (${Theme})::$totals - lista w logu kroku i w ui-preview/$($Theme.ToLower())/uwagi-motyw.txt"
    }
    if (-not [string]::IsNullOrWhiteSpace($env:GITHUB_STEP_SUMMARY)) {
        try {
            $md = New-Object System.Text.StringBuilder
            [void]$md.AppendLine("### Podgląd UI - motyw $Theme")
            [void]$md.AppendLine('')
            [void]$md.AppendLine("PNG: $global:UiShotIndex | $totals")
            $details = New-Object System.Collections.Generic.List[string]
            foreach ($line in $summary) { if ($line -match '^(CLIP|STALE|GLYPH|EMOJI|CONTRAST|ERR) ') { $details.Add($line) } }
            if ($details.Count -gt 0) {
                [void]$md.AppendLine('')
                [void]$md.AppendLine('```text')
                $shown = 0
                foreach ($line in $details) {
                    if ($shown -ge 300) { [void]$md.AppendLine("... i $($details.Count - $shown) kolejnych (pełna lista w artefakcie)"); break }
                    [void]$md.AppendLine($line)
                    $shown++
                }
                [void]$md.AppendLine('```')
            }
            [void]$md.AppendLine('')
            [System.IO.File]::AppendAllText($env:GITHUB_STEP_SUMMARY, $md.ToString(), (New-Object System.Text.UTF8Encoding($false)))
        } catch { Write-Host "Nie udało się zapisać podsumowania zadania: $($_.Exception.Message)" }
    }
}

if ($FailOnClipping -and $global:UiClipping.Count -gt 0) { exit 1 }
if ($FailOnFindings -and ($global:UiFindings.Count + $global:UiErrors.Count) -gt 0) { exit 1 }
