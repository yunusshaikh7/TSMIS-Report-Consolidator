"""Single source of truth for app identity and version.

Imported by the build tooling and the GUI header. Keep this file
dependency-free so it can be imported from anywhere, including the .spec.
"""

__version__ = "0.1.0"             # semantic version MAJOR.MINOR.PATCH
APP_NAME = "TSMIS Consolidator"   # onefolder / executable name
