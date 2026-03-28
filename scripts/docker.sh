#!/usr/bin/env bash
set -euo pipefail

IMAGE_NAME="${IMAGE_NAME:-stoptions-analyzer}"
IMAGE_TAG="${IMAGE_TAG:-latest}"
IMAGE_REF="${IMAGE_NAME}:${IMAGE_TAG}"

usage() {
  cat <<USAGE
Usage: ./scripts/docker.sh <command> [args...]

Commands:
  build                 Build the Docker image
  test [pytest args...] Run smoke tests inside the container (default fast sanity)
  test-all [pytest args...]
                        Run the full pytest suite inside the container
  shell                 Open an interactive shell in the container
  run <cmd...>          Run any command in the container

Examples:
  ./scripts/docker.sh build
  ./scripts/docker.sh test
  ./scripts/docker.sh test-all
  ./scripts/docker.sh run pytest -m smoke -ra
  ./scripts/docker.sh shell
USAGE
}

cmd="${1:-}"
if [[ -z "$cmd" ]]; then
  usage
  exit 1
fi
shift || true

case "$cmd" in
  build)
    docker build -t "$IMAGE_REF" .
    ;;
  test)
    docker run --rm "$IMAGE_REF" pytest -q -m smoke "$@"
    ;;
  test-all)
    docker run --rm "$IMAGE_REF" pytest -q "$@"
    ;;
  shell)
    docker run --rm -it \
      -v "$(pwd)":/app \
      -w /app \
      "$IMAGE_REF" bash
    ;;
  run)
    if [[ "$#" -eq 0 ]]; then
      echo "Error: provide a command to run." >&2
      usage
      exit 1
    fi
    docker run --rm \
      -v "$(pwd)":/app \
      -w /app \
      "$IMAGE_REF" "$@"
    ;;
  *)
    echo "Unknown command: $cmd" >&2
    usage
    exit 1
    ;;
esac
