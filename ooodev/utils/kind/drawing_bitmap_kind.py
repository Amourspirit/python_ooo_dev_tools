from enum import Enum


class DrawingBitmapKind(str, Enum):
    """
    Drawing (some) Bitmap Values.
    The actual value may change from version to version and may be different
    in other languages.

    These are just quick suggestion that show up in Draw Bitmap options.
    """

    BATHROOM_TILES = "Bathroom Tiles"
    BRICK_WALL = "Brick Wall"
    CARDBOARD = "Cardboard"
    COFFEE_BEANS = "Coffee Beans"
    COLOR_STRIPES = "Color Stripes"
    COLORFUL_PEBBLES = "Colorful Pebbles"
    CONCRETE = "Concrete"
    FENCE = "Fence"
    FLORAL = "Floral"
    ICE_LIGHT = "Ice light"
    INVOICE_PAPER = "Invoice Paper"
    LAWN = "Lawn"
    LITTLE_CLOUDS = "Little Clouds"
    MAPLE_LEAVES = "Maple Leaves"
    MARBLE = "Marble"
    PAINTED_WHITE = "Painted White"
    PAPER_CRUMPLED = "Paper Crumpled"
    PAPER_GRAPH = "Paper Graph"
    PAPER_TEXTURE = "Paper Texture"
    PARCHMENT_PAPER = "Parchment Paper"
    POOL = "Pool"
    SAND_LIGHT = "Sand light"
    SPACE = "Space"
    STONE_WALL = "Stone Wall"
    STONE = "Stone"
    STUDIO = "Studio"
    SURFACE = "Surface"
    WALL_OF_ROCK = "Wall of Rock"
    WHITE_DIFFUSION = "White Diffusion"
    ZEBRA = "Zebra"

    def __str__(self) -> str:
        return self.value
