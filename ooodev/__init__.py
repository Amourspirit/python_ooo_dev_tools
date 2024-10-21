import uno  # noqa # type: ignore

# importing using a text file as version source is not supported when packaging scripts using oooscript.
# with open(os.path.join(os.path.dirname(__file__), "VERSION"), "r", encoding="utf-8") as f:
#     version = f.read().strip()

__version__ = "0.49.0"


def get_version() -> str:
    return __version__
