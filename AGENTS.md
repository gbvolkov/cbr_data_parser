# Working Rules

These rules apply to all future work in this repository.

The project is `uv`-based. Use `uv` for environment and dependency management.

1. Keep every solution simple and explicit. Do not introduce unnecessary abstractions, classes, or indirection.
2. Do not add fallbacks or workarounds. If the environment is broken or the input data is invalid, report the error and stop.
3. Preserve layered architecture. Keep scenario code separate from utility code.
4. Prefer UTF-8 whenever possible. Do not use Unicode escape sequences such as `\uXXXX` unless there is a hard technical requirement.
