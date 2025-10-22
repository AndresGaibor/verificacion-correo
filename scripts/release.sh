#!/bin/bash
#
# Release Automático - Verificación de Correos OWA
#
# Crea y sube automáticamente un nuevo tag de versión siguiendo semántica vX.Y.Z
# Esto dispara el workflow de GitHub Actions para build y release
#
# Uso:
#   ./scripts/release.sh [patch|minor|major] [opciones]
#
# Argumentos:
#   patch       Incrementa versión patch (v1.0.0 → v1.0.1) [por defecto]
#   minor       Incrementa versión minor (v1.0.0 → v1.1.0)
#   major       Incrementa versión major (v1.0.0 → v2.0.0)
#
# Opciones:
#   --force         No pedir confirmación
#   --dry-run       Mostrar qué haría sin ejecutar
#   --allow-dirty   Permitir cambios sin commitear
#   -m, --message   Mensaje personalizado para el tag
#
# Ejemplos:
#   ./scripts/release.sh                    # Incrementa patch
#   ./scripts/release.sh minor              # Incrementa minor
#   ./scripts/release.sh major --force      # Incrementa major sin confirmar
#   ./scripts/release.sh patch -m "Bug fix" # Con mensaje personalizado

set -e  # Salir en caso de error

# Colores para output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
CYAN='\033[0;36m'
BOLD='\033[1m'
NC='\033[0m' # No Color

# Variables globales
BUMP_TYPE="patch"
FORCE=false
DRY_RUN=false
ALLOW_DIRTY=false
TAG_MESSAGE=""
REPO_URL="https://github.com/AndresGaibor/verificacion-correo"

# Funciones de utilidad
print_header() {
    echo -e "${BOLD}${CYAN}"
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    echo "$1"
    echo "━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    echo -e "${NC}"
}

print_success() {
    echo -e "${GREEN}✅ $1${NC}"
}

print_error() {
    echo -e "${RED}❌ Error: $1${NC}" >&2
}

print_warning() {
    echo -e "${YELLOW}⚠️  $1${NC}"
}

print_info() {
    echo -e "${BLUE}ℹ️  $1${NC}"
}

print_step() {
    echo -e "${CYAN}▶️  $1${NC}"
}

# Función para mostrar ayuda
show_help() {
    cat << EOF
🚀 Release Automático - Verificación de Correos OWA

Uso: ./scripts/release.sh [TIPO] [OPCIONES]

Tipos de incremento:
  patch       Incrementa versión patch (v1.0.0 → v1.0.1) [por defecto]
  minor       Incrementa versión minor (v1.0.0 → v1.1.0)
  major       Incrementa versión major (v1.0.0 → v2.0.0)

Opciones:
  --force         No pedir confirmación antes de crear el tag
  --dry-run       Mostrar qué haría sin ejecutar comandos
  --allow-dirty   Permitir cambios sin commitear en el repositorio
  -m, --message   Mensaje personalizado para el tag
  -h, --help      Mostrar esta ayuda

Ejemplos:
  ./scripts/release.sh                           # Incrementa patch
  ./scripts/release.sh minor                     # Incrementa minor
  ./scripts/release.sh major --force             # Incrementa major sin confirmar
  ./scripts/release.sh patch -m "Hotfix crítico" # Con mensaje personalizado
  ./scripts/release.sh minor --dry-run           # Ver qué haría sin ejecutar

El script:
  1. Obtiene el último tag del repositorio
  2. Incrementa la versión según el tipo especificado
  3. Crea el tag localmente
  4. Sube el tag a GitHub (origin)
  5. Dispara automáticamente el workflow de build y release

Más información: $REPO_URL
EOF
    exit 0
}

# Parsear argumentos
while [[ $# -gt 0 ]]; do
    case $1 in
        patch|minor|major)
            BUMP_TYPE="$1"
            shift
            ;;
        --force)
            FORCE=true
            shift
            ;;
        --dry-run)
            DRY_RUN=true
            shift
            ;;
        --allow-dirty)
            ALLOW_DIRTY=true
            shift
            ;;
        -m|--message)
            TAG_MESSAGE="$2"
            shift 2
            ;;
        -h|--help)
            show_help
            ;;
        *)
            print_error "Argumento desconocido: $1"
            echo "Usa './scripts/release.sh --help' para ver la ayuda"
            exit 1
            ;;
    esac
done

# Validaciones previas
print_header "🚀 Release Automático - Verificación de Correos OWA"

# 1. Verificar que git esté instalado
if ! command -v git &> /dev/null; then
    print_error "Git no está instalado"
    exit 1
fi

# 2. Verificar que estamos en un repositorio git
if ! git rev-parse --git-dir > /dev/null 2>&1; then
    print_error "No estás en un repositorio git"
    exit 1
fi

# 3. Verificar que el remoto origin existe
if ! git remote get-url origin > /dev/null 2>&1; then
    print_error "No existe el remoto 'origin'"
    exit 1
fi

# 4. Verificar que el directorio de trabajo esté limpio (a menos que --allow-dirty)
if [[ "$ALLOW_DIRTY" == false ]]; then
    if [[ -n $(git status --porcelain) ]]; then
        print_error "El directorio de trabajo tiene cambios sin commitear"
        print_info "Usa --allow-dirty para ignorar esta validación"
        echo ""
        git status --short
        exit 1
    fi
fi

# Obtener el branch actual
CURRENT_BRANCH=$(git rev-parse --abbrev-ref HEAD)

# Obtener el último tag
print_step "Obteniendo último tag del repositorio..."
LAST_TAG=$(git tag --sort=-v:refname | head -n 1)

if [[ -z "$LAST_TAG" ]]; then
    print_warning "No se encontraron tags en el repositorio"
    print_info "Se creará el tag inicial v1.0.0"
    LAST_TAG="v0.0.0"
fi

# Parsear versión actual
if [[ $LAST_TAG =~ ^v([0-9]+)\.([0-9]+)\.([0-9]+)$ ]]; then
    MAJOR="${BASH_REMATCH[1]}"
    MINOR="${BASH_REMATCH[2]}"
    PATCH="${BASH_REMATCH[3]}"
else
    print_error "Formato de tag inválido: $LAST_TAG (se espera vX.Y.Z)"
    exit 1
fi

# Calcular nueva versión
case "$BUMP_TYPE" in
    major)
        NEW_MAJOR=$((MAJOR + 1))
        NEW_MINOR=0
        NEW_PATCH=0
        ;;
    minor)
        NEW_MAJOR=$MAJOR
        NEW_MINOR=$((MINOR + 1))
        NEW_PATCH=0
        ;;
    patch)
        NEW_MAJOR=$MAJOR
        NEW_MINOR=$MINOR
        NEW_PATCH=$((PATCH + 1))
        ;;
esac

NEW_TAG="v${NEW_MAJOR}.${NEW_MINOR}.${NEW_PATCH}"

# Verificar que el tag no exista ya
if git rev-parse "$NEW_TAG" >/dev/null 2>&1; then
    print_error "El tag $NEW_TAG ya existe"
    exit 1
fi

# Mostrar información
echo ""
echo -e "${BOLD}📌 Último tag:${NC}     $LAST_TAG"
echo -e "${BOLD}📈 Nueva versión:${NC}  $NEW_TAG (${BUMP_TYPE})"
echo -e "${BOLD}🌿 Branch:${NC}         $CURRENT_BRANCH"
if [[ -n "$TAG_MESSAGE" ]]; then
    echo -e "${BOLD}💬 Mensaje:${NC}        $TAG_MESSAGE"
fi
echo ""

# Si es dry-run, solo mostrar y salir
if [[ "$DRY_RUN" == true ]]; then
    print_info "Modo DRY-RUN - No se ejecutarán comandos"
    echo ""
    echo "Comandos que se ejecutarían:"
    if [[ -n "$TAG_MESSAGE" ]]; then
        echo "  git tag -a \"$NEW_TAG\" -m \"$TAG_MESSAGE\""
    else
        echo "  git tag -a \"$NEW_TAG\" -m \"Release $NEW_TAG\""
    fi
    echo "  git push origin \"$NEW_TAG\""
    echo ""
    print_success "Dry-run completado"
    exit 0
fi

# Pedir confirmación (a menos que --force)
if [[ "$FORCE" == false ]]; then
    echo -e "${YELLOW}¿Crear y subir el tag $NEW_TAG? (s/N)${NC} "
    read -r CONFIRM
    if [[ ! "$CONFIRM" =~ ^[sS]$ ]]; then
        print_warning "Operación cancelada por el usuario"
        exit 0
    fi
fi

echo ""
print_header "Creando Release"

# Crear el tag localmente
print_step "Creando tag local $NEW_TAG..."
if [[ -n "$TAG_MESSAGE" ]]; then
    git tag -a "$NEW_TAG" -m "$TAG_MESSAGE"
else
    git tag -a "$NEW_TAG" -m "Release $NEW_TAG"
fi
print_success "Tag $NEW_TAG creado localmente"

# Subir el tag al remoto
print_step "Subiendo tag a origin..."
if git push origin "$NEW_TAG"; then
    print_success "Tag $NEW_TAG subido a origin"
else
    print_error "Error al subir el tag a origin"
    print_warning "Eliminando tag local..."
    git tag -d "$NEW_TAG"
    exit 1
fi

# Mostrar resultado
echo ""
print_header "🎉 Release Iniciado Exitosamente"
echo ""
echo -e "${GREEN}${BOLD}✨ Tag $NEW_TAG creado y subido${NC}"
echo ""
echo "El workflow de GitHub Actions se ejecutará automáticamente para:"
echo "  • Compilar el ejecutable Windows"
echo "  • Crear el release en GitHub"
echo "  • Generar release notes automáticamente"
echo ""
echo -e "${BOLD}Enlaces útiles:${NC}"
echo -e "  📦 GitHub Actions: ${BLUE}${REPO_URL}/actions${NC}"
echo -e "  📋 Releases:       ${BLUE}${REPO_URL}/releases${NC}"
echo -e "  🏷️  Tags:           ${BLUE}${REPO_URL}/tags${NC}"
echo ""
print_success "¡Listo! Revisa GitHub Actions para ver el progreso del build"