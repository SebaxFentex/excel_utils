# PLAN.md — excel (libreria interna)

Auditoria 2026-11-02: 6 de 29 cumplen, 11 no cumplen, 7 parciales.
Repo local init 2026-11-02 (rama `main`). Remoto pendiente (D3).

## Decisiones (2026-11-02, del usuario)

- **D1**: renombre total de la API a español. El consumidor (`Archivos Cierre`)
  sera redisenado: no se migra aqui.
- **D2**: tests locales sin configuracion externa. Sin `.env`; los `.bat`
  de tests (cuando existan) van sin `--env-file`.
- **D3**: crear remoto `https://github.com/SebaxFentex/excel_utils`.
- **D4**: publicar la version actual `0.1.3` (tag `v0.1.3`, sin bump).
- **D5**: aplicada opcion C (2026-11-02) en version `0.2.0`.
- **Tests**: `tests/` + `ejecutables/tests/` creados (2026-11-02, 3 tests
  locales); tabla dinamica positiva pendiente de fixture.

## Fases

1. **Base versionable** (hecha): `.gitignore` minimo + commit inicial.
2. **Paquete**: `excel/` → `src/excel/`, `pyproject.toml` a `uv_build`,
   `requires-python>=3.14`, pins `==`, `uv lock && uv sync`, `uv build`
   sin warnings, `import excel` smoke.
3. **Codigo**: `_log.py` + `NullHandler`, `_excepciones.py` con
   `codigo_salida`, renombre API a español, `README.md` (7 bloques),
   `AGENTS.md` con historial. Mapeo viejo→nuevo en AGENTS.
4. **Ejecutables** (sin tests): `build.bat`, `abrir_vscode.bat` de
   plantilla; eliminar `abrir vscode.bat` de la raiz.
5. **Publicacion**: crear repo `excel_utils`, `push`, tag `v0.1.3`,
   verificar `ls-remote`. Release: preguntar (no se crea por defecto).

## Aceptacion global

`uv build` sin warnings · `import excel` con entorno vacio (NullHandler,
sin engines) · greps en cero (`getenv`, `print(`, `basicConfig`, `cwd(`,
`-- python`) · `.bat` probados ejecutandolos · tag visible en el remoto.
