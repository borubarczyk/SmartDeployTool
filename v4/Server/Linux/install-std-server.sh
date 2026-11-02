#!/usr/bin/env bash
# STD - Smart Tool for Deployment: instalka serwera plików (Linux + nginx).
#
# Pierwsze uruchomienie instaluje serwer, każde kolejne go aktualizuje. Ustawienia, konta,
# opublikowane konfiguracje i pliki instalatorów zostają. Parametry podane w wierszu
# poleceń nadpisują zapisane ustawienia i też są zapisywane.
#
#   sudo ./install-std-server.sh [--domain NAZWA] [--cert PLIK --key PLIK | --self-signed]
#                                [--port 443] [--data-dir /srv/std] [--yes]
#
# Opis: README.md obok tego pliku. Umowa: v4/SPEC.md, sekcja 5.
set -euo pipefail
shopt -u patsub_replacement 2>/dev/null || true
umask 022
export PATH="/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin${PATH:+:$PATH}"
# Stałe reguły dopasowania znaków ([A-Za-z] tylko ASCII); polskie komunikaty to zwykłe bajty UTF-8.
export LC_ALL=C

readonly STD_SERVER_VERSION='4.0.0'

readonly ETC_DIR=/etc/std-server
readonly SETTINGS_FILE=$ETC_DIR/settings.conf
readonly TLS_DIR=$ETC_DIR/tls
readonly SELF_SIGNED_CRT=$TLS_DIR/std-server.crt
readonly SELF_SIGNED_KEY=$TLS_DIR/std-server.key
readonly TECH_FILE=$ETC_DIR/technicians.htpasswd
readonly ADMIN_FILE=$ETC_DIR/admins.htpasswd
readonly READERS_FILE=$ETC_DIR/readers.htpasswd
readonly BOOTSTRAP_FILE=$ETC_DIR/bootstrap.json
readonly STATE_DIR=/var/lib/std-server
readonly BACKUP_ROOT=$STATE_DIR/backup
readonly ACME_DIR=$STATE_DIR/acme
readonly LOCK_FILE=$STATE_DIR/lock
readonly NGINX_DIR=/etc/nginx
readonly NGINX_MAIN_CONF=$NGINX_DIR/nginx.conf
readonly NGINX_CONF=$NGINX_DIR/conf.d/std-server.conf
readonly NGINX_CONF_NAME=std-server.conf
readonly TOOL_TARGET=/usr/local/sbin/std-server
readonly DEFAULT_PORT=443
readonly DEFAULT_DATA_DIR=/srv/std

SCRIPT_DIR=$(cd -- "$(dirname -- "${BASH_SOURCE[0]}")" && pwd)
readonly SCRIPT_DIR
readonly TOOL_SOURCE=$SCRIPT_DIR/std-server

# --------------------------------------------------------------------------------------------
# Komunikaty
# --------------------------------------------------------------------------------------------

info() { printf '==> %s\n' "$*"; }
note() { printf '    %s\n' "$*"; }
warn() { printf 'UWAGA: %s\n' "$*" >&2; }
die() { printf 'BŁĄD: %s\n' "$*" >&2; exit 1; }
usage_error() { printf 'BŁĄD: %s\n' "$*" >&2; printf 'Pomoc: %s --help\n' "$0" >&2; exit 2; }

usage() {
    cat <<'EOF'
STD - Smart Tool for Deployment: instalka serwera plików (Linux + nginx)

Użycie:
  sudo ./install-std-server.sh [--domain NAZWA] [--cert PLIK --key PLIK | --self-signed]
                               [--port 443] [--data-dir /srv/std] [--yes]

Parametry:
  --domain NAZWA      nazwa serwera w adresie, np. pobieranie.firma.pl
  --cert PLIK         certyfikat TLS w formacie PEM (najlepiej pełny łańcuch, np. fullchain.pem)
  --key PLIK          klucz prywatny do certyfikatu (PEM, bez hasła)
  --self-signed       certyfikat samopodpisany w /etc/std-server/tls/ - tylko do testów
  --port PORT         port HTTPS (domyślnie 443); port 80 zawsze przekierowuje na HTTPS
  --data-dir KATALOG  katalog danych (domyślnie /srv/std); pliki są w KATALOG/std/
  --yes               bez pytań (do skryptów i CI)
  -h, --help          ta pomoc

Aktualizacja: uruchom instalkę ponownie (także bez parametrów). Zapisane ustawienia są
w /etc/std-server/settings.conf. Konta: std-server add-tech | add-admin | remove | list.
EOF
}

# --------------------------------------------------------------------------------------------
# Pliki tymczasowe i blokada
# --------------------------------------------------------------------------------------------

TMP_FILES=()
LOCK_HELD=0
cleanup() {
    local f
    for f in "${TMP_FILES[@]+"${TMP_FILES[@]}"}"; do
        [ -n "$f" ] && rm -f -- "$f"
    done
    if [ "$LOCK_HELD" = 1 ]; then
        flock -u 9 2>/dev/null || true
    fi
}
trap cleanup EXIT

TMP=''
make_temp() {
    # $1 = katalog, $2 = przedrostek. Ustawia zmienną TMP. Plik tymczasowy leży w tym samym
    # katalogu co plik docelowy, żeby podmiana przez mv była atomowa. Funkcji nie wołamy
    # w $(...), bo wtedy lista TMP_FILES (sprzątanie przy wyjściu) nie zostałaby uzupełniona.
    TMP=$(mktemp "$1/.$2.XXXXXX")
    TMP_FILES+=("$TMP")
}

acquire_lock() {
    command -v flock >/dev/null 2>&1 || { warn "Brak polecenia flock - pomijam blokadę przed równoległym uruchomieniem."; return 0; }
    exec 9>"$LOCK_FILE"
    flock -w 60 9 || die "Inna instalka albo std-server działa w tej chwili (blokada $LOCK_FILE). Spróbuj za chwilę."
    LOCK_HELD=1
}
release_lock() {
    if [ "$LOCK_HELD" = 1 ]; then
        flock -u 9 || true
        exec 9>&-
        LOCK_HELD=0
    fi
}

# --------------------------------------------------------------------------------------------
# Walidacja wartości (to one trafiają do konfiguracji nginx, więc są jedyną ochroną
# przed wstrzyknięciem dyrektyw)
# --------------------------------------------------------------------------------------------

is_ipv4() {
    local ip=$1 part
    [[ $ip =~ ^[0-9]{1,3}(\.[0-9]{1,3}){3}$ ]] || return 1
    local IFS=.
    for part in $ip; do
        [ "$((10#$part))" -le 255 ] || return 1
    done
    return 0
}

valid_domain() {
    local d=$1
    is_ipv4 "$d" && return 0
    [ "${#d}" -le 253 ] || return 1
    [[ $d =~ ^[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?(\.[A-Za-z0-9]([A-Za-z0-9-]{0,61}[A-Za-z0-9])?)*$ ]]
}

valid_port() {
    [[ $1 =~ ^[0-9]{1,5}$ ]] || return 1
    [ "$((10#$1))" -ge 1 ] && [ "$((10#$1))" -le 65535 ] && [ "$((10#$1))" -ne 80 ]
}

valid_abs_path() {
    # Ścieżka bezwzględna bez spacji, cudzysłowów, średników i segmentów . / ..
    local p=$1
    [[ $p =~ ^/[A-Za-z0-9._@+/-]+$ ]] || return 1
    case "$p/" in
        *//*|*/./*|*/../*) return 1 ;;
    esac
    return 0
}

valid_data_dir() {
    local p=${1%/}
    valid_abs_path "$p" || return 1
    case $p in
        /|/bin|/boot|/dev|/etc|/home|/lib|/lib32|/lib64|/opt|/proc|/root|/run|/sbin|/srv|/sys|/tmp|/usr|/var|/var/lib|/var/www|/etc/*|/proc/*|/sys/*|/dev/*|/run/*|/usr/*|/boot/*|"$STATE_DIR"|"$STATE_DIR"/*)
            return 1 ;;
    esac
    return 0
}

valid_iso_date() {
    [[ $1 =~ ^[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}:[0-9]{2}Z$ ]]
}

valid_version() {
    [[ $1 =~ ^[0-9]+(\.[0-9]+){1,3}$ ]]
}

valid_unix_name() {
    [[ $1 =~ ^[a-z_][a-z0-9_-]{0,31}$ ]]
}

abs_path() {
    # Zamienia ścieżkę względną na bezwzględną bez rozwijania dowiązań (ważne dla
    # /etc/letsencrypt/live/..., gdzie dowiązania zmieniają się przy odnowieniu).
    local p=$1
    if command -v realpath >/dev/null 2>&1; then
        realpath -s -m -- "$p"
    else
        case $p in
            /*) printf '%s\n' "$p" ;;
            *) printf '%s/%s\n' "$PWD" "${p#./}" ;;
        esac
    fi
}

# --------------------------------------------------------------------------------------------
# Ustawienia: /etc/std-server/settings.conf (KLUCZ=wartość). Plik jest CZYTANY PARSEREM,
# nigdy przez "source" - nic z niego nie jest wykonywane.
# --------------------------------------------------------------------------------------------

declare -A SETTINGS=()
readonly SETTINGS_KEYS='DOMAIN PORT DATA_DIR TLS_MODE CERT_FILE KEY_FILE NGINX_USER NGINX_GROUP INSTALLED_VERSION INSTALLED_AT UPDATED_AT'

load_settings() {
    local file=$1 line key value n=0
    [ -f "$file" ] || return 0
    while IFS= read -r line || [ -n "$line" ]; do
        n=$((n + 1))
        line=${line%$'\r'}
        [[ $line =~ ^[[:space:]]*(#.*)?$ ]] && continue
        if [[ $line =~ ^[[:space:]]*([A-Z][A-Z0-9_]*)[[:space:]]*=[[:space:]]*(.*)$ ]]; then
            key=${BASH_REMATCH[1]}
            value=${BASH_REMATCH[2]}
            value=${value%"${value##*[![:space:]]}"}
            if [[ $value =~ ^\"(.*)\"$ ]] || [[ $value =~ ^\'(.*)\'$ ]]; then
                value=${BASH_REMATCH[1]}
            fi
            case " $SETTINGS_KEYS " in
                *" $key "*) SETTINGS[$key]=$value ;;
                *) warn "$file, linia $n: nieznany klucz $key - pomijam." ;;
            esac
        else
            warn "$file, linia $n: niezrozumiała linia - pomijam."
        fi
    done < "$file"
}

setting() {
    # $1 = klucz; pusty tekst, gdy brak
    printf '%s' "${SETTINGS[$1]-}"
}

# --------------------------------------------------------------------------------------------
# Parametry
# --------------------------------------------------------------------------------------------

ARG_DOMAIN=''
ARG_CERT=''
ARG_KEY=''
ARG_SELF_SIGNED=0
ARG_PORT=''
ARG_DATA_DIR=''
ASSUME_YES=0

need_value() {
    if [ "$#" -lt 2 ] || [ -z "$2" ]; then
        usage_error "Parametr $1 wymaga wartości."
    fi
    case $2 in
        --*) usage_error "Parametr $1 wymaga wartości (podano $2)." ;;
    esac
}

parse_args() {
    while [ "$#" -gt 0 ]; do
        case $1 in
            --domain) need_value "$@"; ARG_DOMAIN=$2; shift 2 ;;
            --domain=*) ARG_DOMAIN=${1#*=}; shift ;;
            --cert) need_value "$@"; ARG_CERT=$2; shift 2 ;;
            --cert=*) ARG_CERT=${1#*=}; shift ;;
            --key) need_value "$@"; ARG_KEY=$2; shift 2 ;;
            --key=*) ARG_KEY=${1#*=}; shift ;;
            --self-signed) ARG_SELF_SIGNED=1; shift ;;
            --port) need_value "$@"; ARG_PORT=$2; shift 2 ;;
            --port=*) ARG_PORT=${1#*=}; shift ;;
            --data-dir) need_value "$@"; ARG_DATA_DIR=$2; shift 2 ;;
            --data-dir=*) ARG_DATA_DIR=${1#*=}; shift ;;
            -y|--yes) ASSUME_YES=1; shift ;;
            -h|--help) usage; exit 0 ;;
            *) usage_error "Nieznany parametr: $1" ;;
        esac
    done
    if [ "$ARG_SELF_SIGNED" = 1 ] && { [ -n "$ARG_CERT" ] || [ -n "$ARG_KEY" ]; }; then
        usage_error "Podaj albo --cert i --key, albo --self-signed - nie oba naraz."
    fi
    if { [ -n "$ARG_CERT" ] && [ -z "$ARG_KEY" ]; } || { [ -z "$ARG_CERT" ] && [ -n "$ARG_KEY" ]; }; then
        usage_error "--cert i --key podaje się razem."
    fi
}

# --------------------------------------------------------------------------------------------
# System: pakiety, nginx, systemd
# --------------------------------------------------------------------------------------------

have_systemd() {
    [ -d /run/systemd/system ] && command -v systemctl >/dev/null 2>&1
}

PKG_MANAGER=''
detect_package_manager() {
    if command -v apt-get >/dev/null 2>&1; then
        PKG_MANAGER=apt
    elif command -v dnf >/dev/null 2>&1; then
        PKG_MANAGER=dnf
    elif command -v yum >/dev/null 2>&1; then
        PKG_MANAGER=yum
    else
        PKG_MANAGER=''
    fi
}

selinux_enabled() {
    command -v selinuxenabled >/dev/null 2>&1 && selinuxenabled
}

ensure_packages() {
    local missing=() pkgs=() p
    command -v nginx >/dev/null 2>&1 || missing+=(nginx)
    command -v htpasswd >/dev/null 2>&1 || missing+=(htpasswd)
    command -v openssl >/dev/null 2>&1 || missing+=(openssl)
    if [ "${#missing[@]}" -eq 0 ]; then
        note "nginx, htpasswd i openssl są już zainstalowane."
        return 0
    fi
    case $PKG_MANAGER in
        apt)
            for p in "${missing[@]}"; do
                case $p in
                    htpasswd) pkgs+=(apache2-utils) ;;
                    *) pkgs+=("$p") ;;
                esac
            done
            info "Instaluję pakiety: ${pkgs[*]}"
            # 9>&- : usługi uruchamiane przez skrypty pakietów nie mogą dziedziczyć blokady instalki.
            DEBIAN_FRONTEND=noninteractive apt-get update -q 9>&- || die "apt-get update nie powiódł się."
            DEBIAN_FRONTEND=noninteractive apt-get install -y -q "${pkgs[@]}" 9>&- || die "Instalacja pakietów nie powiodła się: ${pkgs[*]}"
            ;;
        dnf|yum)
            for p in "${missing[@]}"; do
                case $p in
                    htpasswd) pkgs+=(httpd-tools) ;;
                    *) pkgs+=("$p") ;;
                esac
            done
            info "Instaluję pakiety: ${pkgs[*]}"
            "$PKG_MANAGER" install -y "${pkgs[@]}" 9>&- || die "Instalacja pakietów nie powiodła się: ${pkgs[*]}"
            ;;
        *)
            die "Brak programów: ${missing[*]}. Ten system nie ma apt-get ani dnf - zainstaluj nginx, htpasswd i openssl ręcznie."
            ;;
    esac
    for p in nginx htpasswd openssl; do
        command -v "$p" >/dev/null 2>&1 || die "Po instalacji pakietów nadal brak programu $p."
    done
}

check_nginx_build() {
    local v
    v=$(nginx -V 2>&1 || true)
    case $v in
        *--with-http_dav_module*) ;;
        *) die "Ta wersja nginx nie ma modułu WebDAV (http_dav_module), potrzebnego do zapisu do /std/drop/ i /std-admin/. Zainstaluj pełny pakiet nginx z repozytorium systemu." ;;
    esac
    case $v in
        *--with-http_ssl_module*) ;;
        *) die "Ta wersja nginx nie ma modułu TLS (http_ssl_module)." ;;
    esac
    [ -f "$NGINX_MAIN_CONF" ] || die "Brak pliku $NGINX_MAIN_CONF."
    if ! grep -Eq '^[[:space:]]*include[[:space:]]+[^;#]*conf\.d/\*\.conf[[:space:]]*;' "$NGINX_MAIN_CONF"; then
        die "$NGINX_MAIN_CONF nie wczytuje plików conf.d/*.conf (dyrektywa include w sekcji http). Dodaj ją i uruchom instalkę ponownie."
    fi
}

detect_nginx_user() {
    local u
    u=$(awk '!/^[[:space:]]*#/ && $1 == "user" { sub(/;.*/, "", $2); print $2; exit }' "$NGINX_MAIN_CONF" 2>/dev/null || true)
    if [ -n "$u" ] && valid_unix_name "$u" && id -u "$u" >/dev/null 2>&1; then
        printf '%s\n' "$u"
        return 0
    fi
    for u in www-data nginx; do
        if id -u "$u" >/dev/null 2>&1; then
            printf '%s\n' "$u"
            return 0
        fi
    done
    return 1
}

nginx_pid_file() {
    local p
    p=$(awk '!/^[[:space:]]*#/ && $1 == "pid" { sub(/;.*/, "", $2); print $2; exit }' "$NGINX_MAIN_CONF" 2>/dev/null || true)
    if [ -z "$p" ]; then
        p=$(nginx -V 2>&1 | tr ' ' '\n' | sed -n 's/^--pid-path=//p' | head -n 1 || true)
    fi
    printf '%s\n' "${p:-/run/nginx.pid}"
}

nginx_master_running() {
    local pf pid
    pf=$(nginx_pid_file)
    [ -r "$pf" ] || return 1
    pid=$(head -n 1 "$pf" 2>/dev/null || true)
    [[ $pid =~ ^[0-9]+$ ]] && kill -0 "$pid" 2>/dev/null
}

nginx_running() {
    if have_systemd && systemctl is-active --quiet nginx; then
        return 0
    fi
    nginx_master_running
}

nginx_signal_reload() {
    # "nginx -s reload" wypisuje komunikat [notice] na stderr - pokazujemy go tylko przy błędzie.
    local out
    if ! out=$(nginx -s reload 2>&1 9>&-); then
        printf '%s\n' "$out" >&2
        return 1
    fi
}

reload_or_start_nginx() {
    # Bez systemd (np. kontener) nginx jest uruchamiany i przeładowywany bezpośrednio.
    if have_systemd; then
        systemctl enable nginx >/dev/null 2>&1 || warn "Nie udało się włączyć nginx przy starcie systemu (systemctl enable nginx)."
        if systemctl is-active --quiet nginx; then
            note "Przeładowanie nginx (systemctl reload nginx)."
            systemctl reload nginx 9>&-
        elif nginx_master_running; then
            note "nginx działa poza systemd - przeładowanie przez nginx -s reload."
            nginx_signal_reload
        else
            note "Uruchamiam nginx (systemctl start nginx)."
            systemctl start nginx 9>&-
        fi
    else
        if nginx_master_running; then
            note "Brak systemd - przeładowanie przez nginx -s reload."
            nginx_signal_reload
        else
            note "Brak systemd - uruchamiam nginx bezpośrednio."
            # 9>&- : demon nginx nie może dziedziczyć deskryptora blokady instalki.
            nginx 9>&-
        fi
    fi
}

# --------------------------------------------------------------------------------------------
# Kopie zapasowe
# --------------------------------------------------------------------------------------------

BACKUP_DIR=''
backup_file() {
    # $1 = plik; kopia trafia do /var/lib/std-server/backup/<data>/ (jeden katalog na uruchomienie)
    [ -f "$1" ] || return 0
    if [ -z "$BACKUP_DIR" ]; then
        BACKUP_DIR=$BACKUP_ROOT/$(date +%Y%m%d-%H%M%S)
        if [ -e "$BACKUP_DIR" ]; then
            BACKUP_DIR=$BACKUP_DIR-$$
        fi
        mkdir -p -- "$BACKUP_ROOT"
        chmod 700 -- "$BACKUP_ROOT"
        mkdir -m 700 -- "$BACKUP_DIR"
    fi
    cp -p -- "$1" "$BACKUP_DIR/"
}

# --------------------------------------------------------------------------------------------
# Katalogi, pliki haseł, uprawnienia
# --------------------------------------------------------------------------------------------

prepare_directories() {
    install -d -m 755 -o root -g root "$ETC_DIR"
    install -d -m 700 -o root -g root "$TLS_DIR"
    install -d -m 755 -o root -g root "$STATE_DIR"
    install -d -m 700 -o root -g root "$BACKUP_ROOT"
    install -d -m 755 -o root -g root "$ACME_DIR"
    install -d -m 755 -o root -g root "$ACME_DIR/.well-known"
    install -d -m 755 -o root -g root "$ACME_DIR/.well-known/acme-challenge"
}

prepare_data_dir() {
    local d
    if [ ! -d "$DATA_DIR" ]; then
        install -d -m 755 -o root -g root "$DATA_DIR"
    fi
    for d in std std/clients std/files std/tool std/drop; do
        if [ ! -d "$DATA_DIR/$d" ]; then
            mkdir -p -- "$DATA_DIR/$d"
        fi
    done
    if [ ! -d "$DATA_DIR/upload-tmp" ]; then
        mkdir -p -- "$DATA_DIR/upload-tmp"
    fi
    chown "$NGINX_USER:$NGINX_GROUP" -- "$DATA_DIR/upload-tmp"
    chmod 700 -- "$DATA_DIR/upload-tmp"
    fix_data_permissions
}

fix_data_permissions() {
    # Właściciel: użytkownik nginx; katalogi 750, pliki 640. Dowiązania symboliczne
    # pomijamy (find bez -L, tylko -type d / -type f), więc nic poza katalogiem danych
    # nie zmienia właściciela. Niczego nie usuwamy.
    # -H: sam katalog std może być dowiązaniem (np. na inny dysk), dowiązania w środku nie.
    local root=$DATA_DIR/std
    find -H "$root" \( -type d -o -type f \) ! \( -user "$NGINX_USER" -group "$NGINX_GROUP" \) -exec chown -h "$NGINX_USER:$NGINX_GROUP" {} +
    find -H "$root" -type d ! -perm 750 -exec chmod 750 {} +
    find -H "$root" -type f ! -perm 640 -exec chmod 640 {} +
}

ensure_password_file() {
    # Tworzy pusty plik haseł, jeśli go nie ma. Istniejącego nigdy nie skraca ani nie usuwa.
    local f=$1
    if [ ! -e "$f" ]; then
        install -m 640 -o root -g "$NGINX_GROUP" /dev/null "$f"
        note "Utworzono pusty plik kont $f."
    fi
    chown "root:$NGINX_GROUP" -- "$f"
    chmod 640 -- "$f"
}

rebuild_readers() {
    # readers.htpasswd = administratorzy + technicy (pierwszy wpis danego loginu wygrywa).
    local tmp
    make_temp "$ETC_DIR" readers
    tmp=$TMP
    awk -F: 'NF >= 2 && $1 != "" && $1 !~ /^[[:space:]]*#/ && !seen[$1]++' "$ADMIN_FILE" "$TECH_FILE" > "$tmp"
    chown "root:$NGINX_GROUP" -- "$tmp"
    chmod 640 -- "$tmp"
    mv -f -- "$tmp" "$READERS_FILE"
}

count_accounts() {
    awk -F: 'NF >= 2 && $1 != "" && $1 !~ /^[[:space:]]*#/ { n++ } END { print n + 0 }' "$1"
}

install_tool() {
    [ -f "$TOOL_SOURCE" ] || die "Brak pliku $TOOL_SOURCE (powinien leżeć obok instalki)."
    if [ -f "$TOOL_TARGET" ] && cmp -s "$TOOL_SOURCE" "$TOOL_TARGET"; then
        note "$TOOL_TARGET bez zmian."
        return 0
    fi
    backup_file "$TOOL_TARGET"
    install -d -m 755 /usr/local/sbin
    install -m 755 -o root -g root "$TOOL_SOURCE" "$TOOL_TARGET"
    note "Zainstalowano $TOOL_TARGET."
}

# --------------------------------------------------------------------------------------------
# Certyfikat TLS
# --------------------------------------------------------------------------------------------

cert_matches_host() {
    # $1 = certyfikat, $2 = nazwa / adres IP
    local out
    if is_ipv4 "$2"; then
        out=$(openssl x509 -in "$1" -noout -checkip "$2" 2>/dev/null || true)
    else
        out=$(openssl x509 -in "$1" -noout -checkhost "$2" 2>/dev/null || true)
    fi
    case $out in
        *"does match"*) return 0 ;;
        *) return 1 ;;
    esac
}

public_key_hash_of_cert() {
    openssl x509 -in "$1" -noout -pubkey 2>/dev/null | openssl pkey -pubin -outform DER 2>/dev/null | openssl dgst -sha256 2>/dev/null
}

public_key_hash_of_key() {
    openssl pkey -in "$1" -passin pass: -pubout -outform DER 2>/dev/null </dev/null | openssl dgst -sha256 2>/dev/null
}

self_signed_is_usable() {
    [ -s "$SELF_SIGNED_CRT" ] && [ -s "$SELF_SIGNED_KEY" ] || return 1
    openssl x509 -in "$SELF_SIGNED_CRT" -noout -checkend 2592000 >/dev/null 2>&1 || return 1
    cert_matches_host "$SELF_SIGNED_CRT" "$DOMAIN" || return 1
    local a b
    a=$(public_key_hash_of_cert "$SELF_SIGNED_CRT") || return 1
    b=$(public_key_hash_of_key "$SELF_SIGNED_KEY") || return 1
    [ -n "$a" ] && [ "$a" = "$b" ]
}

generate_self_signed() {
    local san key crt out
    if is_ipv4 "$DOMAIN"; then
        san="IP:$DOMAIN"
    else
        san="DNS:$DOMAIN"
    fi
    make_temp "$TLS_DIR" key
    key=$TMP
    make_temp "$TLS_DIR" crt
    crt=$TMP
    if ! out=$(openssl req -x509 -newkey rsa:3072 -sha256 -days 825 -nodes \
        -keyout "$key" -out "$crt" -subj "/CN=$DOMAIN" \
        -addext "subjectAltName=$san" -addext "extendedKeyUsage=serverAuth" 2>&1 </dev/null); then
        printf '%s\n' "$out" >&2
        die "Nie udało się utworzyć certyfikatu samopodpisanego."
    fi
    chmod 600 -- "$key"
    chmod 644 -- "$crt"
    mv -f -- "$key" "$SELF_SIGNED_KEY"
    mv -f -- "$crt" "$SELF_SIGNED_CRT"
    note "Utworzono certyfikat samopodpisany dla $DOMAIN (825 dni): $SELF_SIGNED_CRT"
}

validate_cert_pair() {
    local crt=$1 key=$2 a b
    if [ ! -f "$crt" ] || [ ! -r "$crt" ]; then
        die "Nie ma pliku certyfikatu $crt (albo nie da się go odczytać)."
    fi
    if [ ! -f "$key" ] || [ ! -r "$key" ]; then
        die "Nie ma pliku klucza $key (albo nie da się go odczytać)."
    fi
    openssl x509 -in "$crt" -noout >/dev/null 2>&1 || die "$crt nie jest certyfikatem w formacie PEM."
    if ! a=$(public_key_hash_of_cert "$crt") || [ -z "$a" ]; then
        die "Nie da się odczytać klucza publicznego z $crt."
    fi
    if ! b=$(public_key_hash_of_key "$key") || [ -z "$b" ]; then
        die "Nie da się odczytać klucza $key. Klucz musi być w formacie PEM i bez hasła."
    fi
    [ "$a" = "$b" ] || die "Klucz $key nie pasuje do certyfikatu $crt."
    if ! openssl x509 -in "$crt" -noout -checkend 0 >/dev/null 2>&1; then
        warn "Certyfikat $crt już wygasł. Komputery z STD odrzucą połączenie."
    elif ! openssl x509 -in "$crt" -noout -checkend 1209600 >/dev/null 2>&1; then
        warn "Certyfikat $crt wygasa w ciągu 14 dni."
    fi
    if ! cert_matches_host "$crt" "$DOMAIN"; then
        warn "Certyfikat $crt nie obejmuje nazwy $DOMAIN. Komputery z STD odrzucą połączenie."
    fi
}

prepare_tls() {
    if [ "$TLS_MODE" = self-signed ]; then
        CERT_FILE=$SELF_SIGNED_CRT
        KEY_FILE=$SELF_SIGNED_KEY
        if self_signed_is_usable; then
            note "Certyfikat samopodpisany $SELF_SIGNED_CRT jest aktualny - zostaje."
        else
            generate_self_signed
        fi
    else
        validate_cert_pair "$CERT_FILE" "$KEY_FILE"
        note "Certyfikat: $CERT_FILE"
    fi
}

# --------------------------------------------------------------------------------------------
# Konfiguracja nginx
# --------------------------------------------------------------------------------------------

ipv6_available() {
    [ -e /proc/net/if_inet6 ]
}

render_nginx_config() {
    local tpl https_origin
    if [ "$PORT" = 443 ]; then
        https_origin="https://$DOMAIN"
    else
        https_origin="https://$DOMAIN:$PORT"
    fi
    # Szablon w apostrofach: zmienne nginx ($uri, $request_uri) zostają bez zmian,
    # a znaczniki @...@ podmieniamy na sprawdzone wcześniej wartości.
    tpl=$(cat <<'NGINX_TEMPLATE'
# STD - Smart Tool for Deployment: serwer plików (nginx).
# Plik generuje install-std-server.sh @VERSION@. Ręczne zmiany zostaną nadpisane przy kolejnym
# uruchomieniu instalki (poprzednia wersja trafia wtedy do /var/lib/std-server/backup/).
# Ustawienia: /etc/std-server/settings.conf. Konta: std-server add-tech | add-admin | remove.

map $uri $std_server_cache_control {
    default    "";
    ~*\.sdt$   "no-cache";
}

# Ochrona przed przypadkowym usunięciem całego drzewa: DELETE katalogu głównego /std-admin/
# i katalogów clients/, files/, tool/, drop/ jest zablokowany (pojedyncze pliki i podkatalogi można usuwać).
map "$request_method:$uri" $std_server_protected {
    default 0;
    "~^DELETE:/std-admin/((clients|files|tool|drop)/?)?$" 1;
}

# HTTP: tylko przekierowanie na HTTPS (oraz pliki weryfikacji Let's Encrypt).
server {
    listen 80;
@IPV6@    listen [::]:80;
    server_name @DOMAIN@;
    server_tokens off;

    location ^~ /.well-known/acme-challenge/ {
        root @ACME_DIR@;
        default_type text/plain;
    }

    location / {
        return 301 @HTTPS_ORIGIN@$request_uri;
    }
}

server {
    listen @PORT@ ssl;
@IPV6@    listen [::]:@PORT@ ssl;
    server_name @DOMAIN@;
    server_tokens off;

    ssl_certificate     @CERT_FILE@;
    ssl_certificate_key @KEY_FILE@;
    ssl_protocols TLSv1.2 TLSv1.3;
    ssl_ciphers ECDHE-ECDSA-AES128-GCM-SHA256:ECDHE-RSA-AES128-GCM-SHA256:ECDHE-ECDSA-AES256-GCM-SHA384:ECDHE-RSA-AES256-GCM-SHA384:ECDHE-ECDSA-CHACHA20-POLY1305:ECDHE-RSA-CHACHA20-POLY1305;
    ssl_prefer_server_ciphers off;
    ssl_session_cache shared:STDSERVER:10m;
    ssl_session_timeout 1d;
    ssl_session_tickets off;

    access_log /var/log/nginx/std-server.access.log;
    error_log  /var/log/nginx/std-server.error.log;

    autoindex off;
    client_max_body_size 1m;
    client_body_temp_path @DATA_DIR@/upload-tmp;

    types {
        application/json         sdt json;
        application/octet-stream msi exe msix zip ps1 xml cab;
    }
    default_type application/octet-stream;

    add_header Strict-Transport-Security "max-age=31536000" always;
    add_header X-Content-Type-Options "nosniff" always;
    add_header Cache-Control $std_server_cache_control always;

    location / {
        return 404;
    }

    # Odczyt: technicy i administratorzy (GET, HEAD).
    location /std/ {
        alias @DATA_DIR@/std/;
        auth_basic "STD";
        auth_basic_user_file @ETC_DIR@/readers.htpasswd;
        limit_except GET HEAD {
            deny all;
        }
    }

    # Raporty i depozyty: tylko zapis (PUT), bez odczytu i listowania, do 50 MB.
    location /std/drop/ {
        alias @DATA_DIR@/std/drop/;
        auth_basic "STD";
        auth_basic_user_file @ETC_DIR@/readers.htpasswd;
        client_max_body_size 50m;
        dav_methods PUT;
        create_full_put_path on;
        dav_access user:rw group:r;
        limit_except PUT {
            deny all;
        }
    }

    # Publikacja: tylko administratorzy, to samo drzewo co /std/, do 4 GB.
    location /std-admin/ {
        alias @DATA_DIR@/std/;
        auth_basic "STD - administracja";
        auth_basic_user_file @ETC_DIR@/admins.htpasswd;
        if ($std_server_protected) {
            return 403;
        }
        client_max_body_size 4g;
        dav_methods PUT DELETE MKCOL;
        create_full_put_path on;
        dav_access user:rw group:r;
        limit_except GET HEAD PUT DELETE MKCOL {
            deny all;
        }
    }
}
NGINX_TEMPLATE
)
    tpl=${tpl//@VERSION@/$STD_SERVER_VERSION}
    tpl=${tpl//@DOMAIN@/$DOMAIN}
    tpl=${tpl//@PORT@/$PORT}
    tpl=${tpl//@HTTPS_ORIGIN@/$https_origin}
    tpl=${tpl//@CERT_FILE@/$CERT_FILE}
    tpl=${tpl//@KEY_FILE@/$KEY_FILE}
    tpl=${tpl//@DATA_DIR@/$DATA_DIR}
    tpl=${tpl//@ETC_DIR@/$ETC_DIR}
    tpl=${tpl//@ACME_DIR@/$ACME_DIR}
    if ipv6_available; then
        printf '%s\n' "$tpl" | sed -e 's/^@IPV6@//'
    else
        printf '%s\n' "$tpl" | sed -e '/^@IPV6@/d'
    fi
}

NGINX_HAD_CONFIG=0
NGINX_CHANGED=0
restore_nginx_config() {
    if [ "$NGINX_HAD_CONFIG" = 1 ] && [ -n "$BACKUP_DIR" ] && [ -f "$BACKUP_DIR/$NGINX_CONF_NAME" ]; then
        cp -p -- "$BACKUP_DIR/$NGINX_CONF_NAME" "$NGINX_CONF"
        note "Przywrócono poprzednią konfigurację z $BACKUP_DIR/$NGINX_CONF_NAME."
    else
        rm -f -- "$NGINX_CONF"
        note "Usunięto nową konfigurację (wcześniej jej nie było)."
    fi
}

apply_nginx_config() {
    local tmp out
    install -d -m 755 "$NGINX_DIR/conf.d"
    # Plik tymczasowy nie kończy się na .conf, więc nginx go nie wczyta.
    make_temp "$NGINX_DIR/conf.d" "$NGINX_CONF_NAME"
    tmp=$TMP
    render_nginx_config > "$tmp"
    chmod 644 -- "$tmp"
    if [ -f "$NGINX_CONF" ] && cmp -s "$tmp" "$NGINX_CONF"; then
        rm -f -- "$tmp"
        note "Konfiguracja nginx bez zmian ($NGINX_CONF)."
        if ! nginx_running; then
            if ! out=$(nginx -t 2>&1); then
                printf '%s\n' "$out" >&2
                die "nginx nie działa, a test konfiguracji (nginx -t) nie przechodzi. Popraw błąd i uruchom instalkę ponownie."
            fi
            reload_or_start_nginx || die "Nie udało się uruchomić nginx. Sprawdź: journalctl -u nginx albo /var/log/nginx/error.log."
        fi
        return 0
    fi

    NGINX_HAD_CONFIG=0
    if [ -f "$NGINX_CONF" ]; then
        backup_file "$NGINX_CONF"
        NGINX_HAD_CONFIG=1
        note "Kopia zapasowa poprzedniej konfiguracji: $BACKUP_DIR/$NGINX_CONF_NAME"
    fi
    mv -f -- "$tmp" "$NGINX_CONF"
    NGINX_CHANGED=1

    if ! out=$(nginx -t 2>&1); then
        printf '%s\n' "$out" >&2
        warn "Test nowej konfiguracji nginx (nginx -t) nie przeszedł."
        restore_nginx_config
        if nginx -t >/dev/null 2>&1; then
            note "Konfiguracja sprzed zmian przechodzi test - serwer działa dalej po staremu."
        else
            warn "Konfiguracja sprzed zmian też nie przechodzi testu, więc błąd nie wynika (tylko) ze zmian STD. Uruchom: nginx -t"
            case $out in
                *"Address family not supported"*)
                    warn "Ten system nie obsługuje IPv6, a inna strona nginx ma 'listen [::]:...' (zwykle /etc/nginx/sites-enabled/default). Usuń te linie albo wyłącz tę stronę i uruchom instalkę ponownie."
                    ;;
            esac
        fi
        die "Nie zmieniono konfiguracji serwera. Ustawienia nie zostały zapisane."
    fi
    note "Test konfiguracji (nginx -t): OK."

    if ! reload_or_start_nginx; then
        warn "Nie udało się przeładować / uruchomić nginx z nową konfiguracją."
        restore_nginx_config
        if nginx -t >/dev/null 2>&1; then
            reload_or_start_nginx || true
        fi
        die "Przywrócono poprzednią konfigurację. Sprawdź, czy port 80 i $PORT nie są zajęte (ss -ltnp) oraz log: journalctl -u nginx albo /var/log/nginx/error.log."
    fi
}

# --------------------------------------------------------------------------------------------
# SELinux i zapora
# --------------------------------------------------------------------------------------------

set_fcontext() {
    # $1 = typ, $2 = wzorzec ścieżki
    semanage fcontext -a -t "$1" "$2" 2>/dev/null || semanage fcontext -m -t "$1" "$2"
}

configure_selinux() {
    selinux_enabled || return 0
    info "SELinux jest włączony - ustawiam etykiety plików."
    if ! command -v semanage >/dev/null 2>&1; then
        case $PKG_MANAGER in
            dnf|yum) "$PKG_MANAGER" install -y policycoreutils-python-utils 9>&- || true ;;
            apt) DEBIAN_FRONTEND=noninteractive apt-get install -y -q policycoreutils-python-utils 9>&- || true ;;
        esac
    fi
    if ! command -v semanage >/dev/null 2>&1; then
        warn "Brak polecenia semanage - ustaw ręcznie etykietę httpd_sys_rw_content_t dla $DATA_DIR/std i $DATA_DIR/upload-tmp."
        return 0
    fi
    set_fcontext httpd_sys_rw_content_t "$DATA_DIR/std(/.*)?" || warn "semanage fcontext dla $DATA_DIR/std nie powiódł się."
    set_fcontext httpd_sys_rw_content_t "$DATA_DIR/upload-tmp(/.*)?" || warn "semanage fcontext dla $DATA_DIR/upload-tmp nie powiódł się."
    set_fcontext httpd_sys_content_t "$ACME_DIR(/.*)?" || warn "semanage fcontext dla $ACME_DIR nie powiódł się."
    restorecon -R -- "$DATA_DIR/std" "$DATA_DIR/upload-tmp" "$ACME_DIR" "$ETC_DIR" || warn "restorecon nie powiódł się."
    if [ "$PORT" != 443 ]; then
        semanage port -a -t http_port_t -p tcp "$PORT" 2>/dev/null || semanage port -m -t http_port_t -p tcp "$PORT" 2>/dev/null || true
    fi
}

configure_firewall() {
    if have_systemd && systemctl is-active --quiet firewalld 2>/dev/null && command -v firewall-cmd >/dev/null 2>&1; then
        info "Zapora firewalld: otwieram port 80 i $PORT/tcp."
        firewall-cmd --permanent --add-service=http >/dev/null || warn "firewall-cmd: nie udało się otworzyć portu 80."
        if [ "$PORT" = 443 ]; then
            firewall-cmd --permanent --add-service=https >/dev/null || warn "firewall-cmd: nie udało się otworzyć portu 443."
        else
            firewall-cmd --permanent --add-port="$PORT/tcp" >/dev/null || warn "firewall-cmd: nie udało się otworzyć portu $PORT."
        fi
        firewall-cmd --reload >/dev/null || warn "firewall-cmd --reload nie powiódł się."
    elif command -v ufw >/dev/null 2>&1 && ufw status 2>/dev/null | grep -q '^Status: active'; then
        info "Zapora ufw: otwieram port 80 i $PORT/tcp."
        ufw allow 80/tcp >/dev/null || warn "ufw: nie udało się otworzyć portu 80."
        ufw allow "$PORT/tcp" >/dev/null || warn "ufw: nie udało się otworzyć portu $PORT."
    fi
}

# --------------------------------------------------------------------------------------------
# Zapis ustawień i pliku bootstrap.json
# --------------------------------------------------------------------------------------------

server_url() {
    if [ "$PORT" = 443 ]; then
        printf 'https://%s/std/\n' "$DOMAIN"
    else
        printf 'https://%s:%s/std/\n' "$DOMAIN" "$PORT"
    fi
}

save_settings() {
    local tmp now
    now=$(date -u +%Y-%m-%dT%H:%M:%SZ)
    [ -n "$INSTALLED_AT" ] || INSTALLED_AT=$now
    make_temp "$ETC_DIR" settings.conf
    tmp=$TMP
    {
        printf '%s\n' '# STD - Smart Tool for Deployment: ustawienia serwera (Linux + nginx).'
        printf '%s\n' '# Plik zapisuje install-std-server.sh. Przy kolejnym uruchomieniu instalka bierze te'
        printf '%s\n' '# wartości jako domyślne; parametry z wiersza poleceń je nadpisują i też są zapisywane.'
        printf '%s\n' '# Format: KLUCZ=wartość, jedna para w linii. Plik jest czytany parserem, nie wykonywany.'
        printf 'DOMAIN=%s\n' "$DOMAIN"
        printf 'PORT=%s\n' "$PORT"
        printf 'DATA_DIR=%s\n' "$DATA_DIR"
        printf 'TLS_MODE=%s\n' "$TLS_MODE"
        printf 'CERT_FILE=%s\n' "$CERT_FILE"
        printf 'KEY_FILE=%s\n' "$KEY_FILE"
        printf 'NGINX_USER=%s\n' "$NGINX_USER"
        printf 'NGINX_GROUP=%s\n' "$NGINX_GROUP"
        printf 'INSTALLED_VERSION=%s\n' "$STD_SERVER_VERSION"
        printf 'INSTALLED_AT=%s\n' "$INSTALLED_AT"
        printf 'UPDATED_AT=%s\n' "$now"
    } > "$tmp"
    chmod 644 -- "$tmp"
    backup_file "$SETTINGS_FILE"
    mv -f -- "$tmp" "$SETTINGS_FILE"
    note "Ustawienia zapisane: $SETTINGS_FILE"
}

write_bootstrap() {
    local tmp
    make_temp "$ETC_DIR" bootstrap.json
    tmp=$TMP
    printf '{\n  "server": "%s",\n  "mirrors": [],\n  "auth": { "mode": "basic" },\n  "proxy": null\n}\n' "$(server_url)" > "$tmp"
    chmod 644 -- "$tmp"
    mv -f -- "$tmp" "$BOOTSTRAP_FILE"
}

# --------------------------------------------------------------------------------------------
# Sprawdzenie po instalacji
# --------------------------------------------------------------------------------------------

SELF_CHECK_CODE=''
probe_server() {
    # Pyta lokalny nginx (127.0.0.1) z właściwą nazwą (SNI), bez serwera pośredniczącego.
    # Oczekiwane: 401 (bez hasła nie ma dostępu).
    local i
    SELF_CHECK_CODE=''
    for i in $(seq 1 20); do
        SELF_CHECK_CODE=$(curl -k -s -o /dev/null -w '%{http_code}' --max-time 5 --noproxy '*' \
            --connect-to "$DOMAIN:$PORT:127.0.0.1:$PORT" "https://$DOMAIN:$PORT/std/clients.sdt" 2>/dev/null || true)
        [ "$SELF_CHECK_CODE" = 401 ] && return 0
        [ "$i" -lt 20 ] && sleep 0.5
    done
    return 1
}

self_check() {
    if ! command -v curl >/dev/null 2>&1; then
        note "Brak curl - pomijam sprawdzenie odpowiedzi serwera."
        return 0
    fi
    if probe_server; then
        note "$(server_url) odpowiada i wymaga logowania (401) - OK."
        return 0
    fi
    warn "Serwer nie odpowiada zgodnie z oczekiwaniem (kod: ${SELF_CHECK_CODE:-brak połączenia}, oczekiwano 401)."
    if [ "$NGINX_CHANGED" = 1 ]; then
        # nginx -t nie wykrywa np. zajętego portu, a "nginx -s reload" zostawia wtedy starą
        # konfigurację w pamięci. Wracamy do poprzedniej, żeby plik i działający serwer się zgadzały.
        restore_nginx_config
        if nginx -t >/dev/null 2>&1; then
            reload_or_start_nginx || true
        fi
        die "Przywrócono poprzednią konfigurację nginx; ustawienia nie zostały zmienione. Sprawdź, czy port $PORT nie jest zajęty (ss -ltnp) i log /var/log/nginx/error.log."
    fi
    warn "Sprawdź: std-server status i /var/log/nginx/std-server.error.log."
}

# --------------------------------------------------------------------------------------------
# Główna procedura
# --------------------------------------------------------------------------------------------

DOMAIN=''
PORT=''
DATA_DIR=''
TLS_MODE=''
CERT_FILE=''
KEY_FILE=''
NGINX_USER=''
NGINX_GROUP=''
INSTALLED_AT=''
PREVIOUS_VERSION=''
INTERACTIVE=0

resolve_settings() {
    local s_domain s_port s_data s_tls s_cert s_key
    s_domain=$(setting DOMAIN)
    s_port=$(setting PORT)
    s_data=$(setting DATA_DIR)
    s_tls=$(setting TLS_MODE)
    s_cert=$(setting CERT_FILE)
    s_key=$(setting KEY_FILE)
    PREVIOUS_VERSION=$(setting INSTALLED_VERSION)
    INSTALLED_AT=$(setting INSTALLED_AT)
    if [ -n "$PREVIOUS_VERSION" ] && ! valid_version "$PREVIOUS_VERSION"; then
        warn "Zapisana wersja ma zły format - pomijam."
        PREVIOUS_VERSION='?'
    fi
    if [ -n "$INSTALLED_AT" ] && ! valid_iso_date "$INSTALLED_AT"; then
        INSTALLED_AT=''
    fi

    # Domena
    if [ -n "$ARG_DOMAIN" ]; then
        DOMAIN=$ARG_DOMAIN
    elif [ -n "$s_domain" ]; then
        DOMAIN=$s_domain
    else
        DOMAIN=$(hostname -f 2>/dev/null || hostname 2>/dev/null || true)
        if [ "$INTERACTIVE" = 1 ]; then
            local answer=''
            read -r -p "Nazwa serwera (domena) [$DOMAIN]: " answer || true
            [ -n "$answer" ] && DOMAIN=$answer
        else
            warn "Nie podano --domain - używam nazwy komputera: ${DOMAIN:-brak}."
        fi
    fi
    DOMAIN=$(printf '%s' "$DOMAIN" | tr '[:upper:]' '[:lower:]')
    valid_domain "$DOMAIN" || usage_error "Nieprawidłowa nazwa serwera: '$DOMAIN'. Dozwolone: litery, cyfry, '-', '.' albo adres IPv4."

    # Port
    if [ -n "$ARG_PORT" ]; then
        PORT=$ARG_PORT
    elif [ -n "$s_port" ]; then
        PORT=$s_port
    else
        PORT=$DEFAULT_PORT
    fi
    valid_port "$PORT" || usage_error "Nieprawidłowy port HTTPS: '$PORT' (1-65535, bez 80 - port 80 służy do przekierowania)."
    PORT=$((10#$PORT))

    # Katalog danych
    if [ -n "$ARG_DATA_DIR" ]; then
        DATA_DIR=$(abs_path "$ARG_DATA_DIR")
    elif [ -n "$s_data" ]; then
        DATA_DIR=$s_data
    else
        DATA_DIR=$DEFAULT_DATA_DIR
    fi
    DATA_DIR=${DATA_DIR%/}
    valid_data_dir "$DATA_DIR" || usage_error "Nieprawidłowy katalog danych: '$DATA_DIR'. Podaj osobny katalog, np. /srv/std (bez spacji i znaków specjalnych, nie katalog systemowy)."
    if [ -n "$s_data" ] && [ "${s_data%/}" != "$DATA_DIR" ]; then
        warn "Zmiana katalogu danych z $s_data na $DATA_DIR. Stary katalog NIE jest przenoszony ani usuwany - przenieś pliki ręcznie, jeśli trzeba."
    fi

    # Certyfikat
    if [ "$ARG_SELF_SIGNED" = 1 ]; then
        TLS_MODE=self-signed
    elif [ -n "$ARG_CERT" ]; then
        TLS_MODE=files
        CERT_FILE=$(abs_path "$ARG_CERT")
        KEY_FILE=$(abs_path "$ARG_KEY")
    elif [ "$s_tls" = self-signed ]; then
        TLS_MODE=self-signed
    elif [ "$s_tls" = files ]; then
        TLS_MODE=files
        CERT_FILE=$s_cert
        KEY_FILE=$s_key
    elif [ -n "$s_tls" ]; then
        usage_error "Zapisany tryb certyfikatu '$s_tls' jest nieznany. Podaj --cert i --key albo --self-signed."
    else
        usage_error "Brak certyfikatu. Podaj --cert PLIK --key PLIK (np. z Let's Encrypt) albo --self-signed (tylko do testów)."
    fi
    if [ "$TLS_MODE" = files ]; then
        valid_abs_path "$CERT_FILE" || usage_error "Nieprawidłowa ścieżka certyfikatu: '$CERT_FILE' (bez spacji i znaków specjalnych)."
        valid_abs_path "$KEY_FILE" || usage_error "Nieprawidłowa ścieżka klucza: '$KEY_FILE' (bez spacji i znaków specjalnych)."
    fi
}

confirm() {
    [ "$INTERACTIVE" = 1 ] || return 0
    local answer=''
    printf '\n'
    if [ -f "$SETTINGS_FILE" ]; then
        printf 'Aktualizacja serwera STD %s -> %s\n' "${PREVIOUS_VERSION:-?}" "$STD_SERVER_VERSION"
    else
        printf 'Instalacja serwera STD %s\n' "$STD_SERVER_VERSION"
    fi
    printf '  Domena:          %s\n' "$DOMAIN"
    printf '  Port HTTPS:      %s\n' "$PORT"
    printf '  Katalog danych:  %s\n' "$DATA_DIR"
    if [ "$TLS_MODE" = self-signed ]; then
        printf '  Certyfikat:      samopodpisany (tylko do testów)\n'
    else
        printf '  Certyfikat:      %s\n' "$CERT_FILE"
        printf '  Klucz:           %s\n' "$KEY_FILE"
    fi
    read -r -p "Kontynuować? [t/N] " answer || true
    case $answer in
        t|T|tak|TAK|Tak|y|Y|yes) ;;
        *) die "Przerwano na życzenie użytkownika. Nic nie zmieniono." ;;
    esac
}

offer_first_admin() {
    [ "$INTERACTIVE" = 1 ] || return 0
    [ "$(count_accounts "$ADMIN_FILE")" -eq 0 ] || return 0
    local login=''
    printf '\nNie ma jeszcze żadnego administratora.\n'
    read -r -p "Login pierwszego administratora (Enter = pomiń): " login || true
    [ -n "$login" ] || return 0
    "$TOOL_TARGET" add-admin "$login" || warn "Nie udało się dodać administratora. Zrób to później: std-server add-admin <login>"
}

main() {
    parse_args "$@"
    [ "$(id -u)" -eq 0 ] || die "Instalkę uruchamia root: sudo $0 $*"
    if [ "$ASSUME_YES" = 0 ]; then
        if [ -t 0 ] && [ -t 1 ]; then
            INTERACTIVE=1
        else
            usage_error "Brak terminala do pytań. Dodaj --yes, żeby instalka działała bez pytań."
        fi
    fi
    [ -f "$TOOL_SOURCE" ] || die "Brak pliku $TOOL_SOURCE (skopiuj cały katalog Server/Linux)."

    info "STD - Smart Tool for Deployment: instalka serwera $STD_SERVER_VERSION (Linux + nginx)"
    if [ -f "$SETTINGS_FILE" ]; then
        load_settings "$SETTINGS_FILE"
        note "Znaleziono ustawienia $SETTINGS_FILE - to jest aktualizacja. Konta i dane zostają."
    else
        note "Pierwsza instalacja."
    fi
    resolve_settings
    confirm

    install -d -m 755 -o root -g root "$STATE_DIR"
    acquire_lock

    detect_package_manager
    info "Pakiety"
    ensure_packages
    check_nginx_build

    NGINX_USER=$(detect_nginx_user) || die "Nie znaleziono użytkownika nginx (www-data / nginx)."
    NGINX_GROUP=$(id -gn "$NGINX_USER") || die "Nie znaleziono grupy użytkownika $NGINX_USER."
    valid_unix_name "$NGINX_GROUP" || die "Nieobsługiwana nazwa grupy nginx: $NGINX_GROUP"
    note "nginx działa jako $NGINX_USER:$NGINX_GROUP."
    if have_systemd; then
        note "systemd: tak."
    else
        note "systemd: nie (nginx będzie uruchamiany i przeładowywany bezpośrednio)."
    fi

    info "Katalogi i uprawnienia"
    prepare_directories
    prepare_data_dir
    note "Dane: $DATA_DIR/std/{clients,files,tool,drop} (właściciel $NGINX_USER, prawa 750 / 640)."

    info "Konta"
    ensure_password_file "$TECH_FILE"
    ensure_password_file "$ADMIN_FILE"
    rebuild_readers
    note "Technicy: $(count_accounts "$TECH_FILE"), administratorzy: $(count_accounts "$ADMIN_FILE"). Pliki kont nie są zmieniane."

    info "Certyfikat TLS"
    prepare_tls

    info "Narzędzie std-server"
    install_tool

    info "Konfiguracja nginx"
    configure_selinux
    apply_nginx_config

    info "Ustawienia"
    save_settings
    write_bootstrap
    configure_firewall

    info "Sprawdzenie"
    self_check

    release_lock

    printf '\n'
    "$TOOL_TARGET" status || true
    offer_first_admin
    if [ "$(count_accounts "$ADMIN_FILE")" -eq 0 ]; then
        printf '\nNastępny krok: dodaj konta (hasło zostanie zapytane):\n'
        printf '  sudo std-server add-admin <login>\n'
        printf '  sudo std-server add-tech <login>\n'
    fi
    if [ -n "$BACKUP_DIR" ]; then
        printf '\nKopia zapasowa poprzednich plików: %s\n' "$BACKUP_DIR"
    fi
    printf '\nGotowe.\n'
}

main "$@"
