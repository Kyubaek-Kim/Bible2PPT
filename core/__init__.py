"""Bible2PPT core: platform-independent logic (no Tkinter imports here).

Submodules:

* :mod:`core.paths`         — OS-independent resource / user-data paths
* :mod:`core.platform_util` — OS-dependent actions (open folder, register font)
* :mod:`core.parser`        — verse-reference parsing/normalisation
* :mod:`core.bible`         — canon, translation registry, reference expansion
* :mod:`core.alignment`     — cross-translation display-layer verse mapping
* :mod:`core.importer`      — user Bible upload parse/validate/register
* :mod:`core.ppt`           — slide engine (pagination, fonts, background)
* :mod:`core.image_util`    — background crop/notify (Pillow)
* :mod:`core.fonts`         — bundled fonts + preview support
* :mod:`core.i18n`          — UI language tables
* :mod:`core.settings`      — settings persistence
* :mod:`core.generator`     — high-level orchestration used by the UI
"""

__all__ = [
    "paths",
    "platform_util",
    "parser",
    "bible",
    "alignment",
    "importer",
    "ppt",
    "image_util",
    "fonts",
    "i18n",
    "settings",
    "generator",
]

__version__ = "1.0.0"
