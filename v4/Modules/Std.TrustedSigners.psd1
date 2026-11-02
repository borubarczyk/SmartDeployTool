# STD (Smart Tool for Deployment) 4 - Copyright (c) 2025-2026 Borys Kaleta. Wszelkie prawa zastrzeżone - warunki w pliku LICENSE.
#
# Zaufane certyfikaty do sprawdzania podpisu konfiguracji STD (SPEC 2.2). Ten plik jest częścią kodu
# i jest podpisywany razem z nim - nie ma innego miejsca (zmienna środowiskowa, plik obok skryptu),
# które mogłoby dodać zaufany certyfikat.
#
# Pusta lista jest dozwolona w repozytorium; wtedy pobranie konfiguracji kończy się błędem
# [STD-SIG-NOTRUST]. Certyfikat do podpisu tworzy Tools\New-StdSigningCertificate.ps1 - wypisuje
# gotowy wpis do wklejenia poniżej.
#
# Thumbprint: odcisk SHA-1 certyfikatu (HEX wielkimi literami), Certificate: base64 DER (plik .cer).
@{
    Signers = @(
        # @{ Name = 'Borys Kaleta - główny'; Thumbprint = 'AB12...'; Certificate = 'MIIC...base64 DER...' }
    )
}
