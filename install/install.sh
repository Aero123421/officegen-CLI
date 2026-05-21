#!/usr/bin/env sh
set -eu

repo="${OFFICEGEN_REPO:-Aero123421/officegen-CLI}"
version="${OFFICEGEN_VERSION:-}"
install_dir="${OFFICEGEN_INSTALL_DIR:-$HOME/.officegen/bin}"

log() {
  printf '%s\n' "$*"
}

fail() {
  printf 'officegen install: %s\n' "$*" >&2
  exit 1
}

need_cmd() {
  command -v "$1" >/dev/null 2>&1 || fail "missing required command: $1"
}

http_get_stdout() {
  if command -v curl >/dev/null 2>&1; then
    curl -fsSL "$1"
  elif command -v wget >/dev/null 2>&1; then
    wget -qO- "$1"
  else
    fail "missing curl or wget"
  fi
}

http_get_file() {
  if command -v curl >/dev/null 2>&1; then
    curl -fL "$1" -o "$2"
  elif command -v wget >/dev/null 2>&1; then
    wget -q "$1" -O "$2"
  else
    fail "missing curl or wget"
  fi
}

detect_target() {
  os="$(uname -s)"
  arch="$(uname -m)"
  case "$os" in
    Linux) os_part="unknown-linux-gnu" ;; # x86_64-unknown-linux-gnu / aarch64-unknown-linux-gnu
    Darwin) os_part="apple-darwin" ;; # x86_64-apple-darwin / aarch64-apple-darwin
    *) fail "unsupported OS: $os" ;;
  esac
  case "$arch" in
    x86_64 | amd64) arch_part="x86_64" ;;
    arm64 | aarch64) arch_part="aarch64" ;;
    *) fail "unsupported CPU architecture: $arch" ;;
  esac
  printf '%s-%s' "$arch_part" "$os_part"
}

resolve_version() {
  if [ -n "$version" ]; then
    printf '%s' "${version#v}"
    return
  fi

  latest_json="$(http_get_stdout "https://api.github.com/repos/$repo/releases/latest")"
  latest_tag="$(printf '%s\n' "$latest_json" | sed -n 's/.*"tag_name":[[:space:]]*"\([^"]*\)".*/\1/p' | head -n 1)"
  [ -n "$latest_tag" ] || fail "could not resolve latest release for $repo"
  printf '%s' "${latest_tag#v}"
}

verify_sha256() {
  archive="$1"
  checksum="$2"
  if command -v sha256sum >/dev/null 2>&1; then
    (cd "$(dirname "$archive")" && sha256sum -c "$(basename "$checksum")")
  elif command -v shasum >/dev/null 2>&1; then
    expected="$(awk '{print $1}' "$checksum")"
    actual="$(shasum -a 256 "$archive" | awk '{print $1}')"
    [ "$expected" = "$actual" ] || fail "checksum mismatch for $(basename "$archive")"
  else
    fail "missing sha256sum or shasum"
  fi
}

target="$(detect_target)"
version="$(resolve_version)"
asset="officegen-v$version-$target.tar.gz"
base_url="https://github.com/$repo/releases/download/v$version"
tmp_dir="$(mktemp -d 2>/dev/null || mktemp -d -t officegen-install)"
cleanup() {
  rm -rf "$tmp_dir"
}
trap cleanup EXIT INT TERM

archive="$tmp_dir/$asset"
checksum="$archive.sha256"
extract_dir="$tmp_dir/extract"

log "Installing officegen v$version for $target"
http_get_file "$base_url/$asset" "$archive"
http_get_file "$base_url/$asset.sha256" "$checksum"
verify_sha256 "$archive" "$checksum"

mkdir -p "$extract_dir" "$install_dir"
tar -xzf "$archive" -C "$extract_dir"
bin_path="$(find "$extract_dir" -type f -name officegen -perm -u+x | head -n 1)"
[ -n "$bin_path" ] || fail "archive did not contain an executable officegen binary"

cp "$bin_path" "$install_dir/officegen"
chmod 755 "$install_dir/officegen"

log "Installed $install_dir/officegen"
export PATH="$install_dir:$(printf '%s' "$PATH" | awk -v RS=: -v ORS=: -v d="$install_dir" '$0 != d { print }' | sed 's/:$//')"

profile_hint=""
case "${SHELL:-}" in
  */zsh) profile_hint="$HOME/.zshrc" ;;
  */bash) profile_hint="$HOME/.bashrc" ;;
  *) profile_hint="$HOME/.profile" ;;
esac

resolved="$(command -v officegen 2>/dev/null || true)"
if [ "$resolved" != "$install_dir/officegen" ]; then
  printf 'officegen install: PATH collision detected. Expected %s but resolved %s\n' "$install_dir/officegen" "${resolved:-<none>}" >&2
  if command -v -a officegen >/dev/null 2>&1; then
    command -v -a officegen >&2 || true
  fi
  fail "put $install_dir before old npm/homebrew/yarn shims in PATH"
fi

actual_version="$(officegen --version)"
[ "$actual_version" = "$version" ] || fail "officegen --version returned $actual_version, expected $version"

case ":$(printf '%s' "${PATH:-}"):" in
  *":$install_dir:"*) ;;
  *) log "Add $install_dir to PATH to run officegen from any directory." ;;
esac

log "officegen v$version is first on PATH for this installer process."
log "To persist this in new shells, ensure this appears before npm/yarn/homebrew shims in $profile_hint:"
log "  export PATH=\"$install_dir:\$PATH\""
