# -*- coding: utf-8 -*-
"""
DEPRECATED MODULE STUB — Do not use.

Ce module est désormais déprécié et ne doit plus être importé.
Veuillez migrer vers la nouvelle architecture :

- tools/server.py : couche MCP (FastMCP) exposant les tools
- utils/* : modules utilitaires spécialisés (pptx/docx/xlsx/pdf, gestion de fichiers,
            upload/download/URL publique, archives, recherche d'images)

Pour référence historique uniquement (lecture seule) :
- tools/file_export_mcp.deprecated.py
"""

from __future__ import annotations

import sys
import warnings

MESSAGE = (
    "tools/file_export_mcp.py est déprécié et n'est plus supporté.\n"
    "Utilisez désormais :\n"
    "  - tools/server.py (couche MCP FastMCP)\n"
    "  - utils/* pour toutes les fonctions utilitaires (pptx/docx/xlsx/pdf, "
    "gestion de fichiers, upload/download/URL publique, archives, recherche d'images)\n"
    "Implémentation historique disponible pour référence : "
    "tools/file_export_mcp.deprecated.py"
)

if __name__ == "__main__":
    print(MESSAGE, file=sys.stderr)
    sys.exit(1)

# Lors d'un import, avertir puis interrompre immédiatement avec une erreur explicite
warnings.warn(MESSAGE, DeprecationWarning, stacklevel=2)
raise ImportError(MESSAGE)
