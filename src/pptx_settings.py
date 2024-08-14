from pathlib import Path
from typing import Union, Tuple
from pptx.util import Inches, Pt
from PIL import ImageColor

Pathlike = Union[str, Path]
BLANK_SLIDE_LAYOUT: int = 6
RGBTriplet = Tuple[int, int, int]
HexColorOrName = str


class PPTXSettings:
    def __init__(self):
        self._top_margin = Inches(1.35)
        self._left_margin = Inches(0.53)
        self._right_margin = Inches(0.53)
        self._bottom_margin = Inches(0.66)
        self._h_center_margin = Inches(0.5)
        self._v_center_margin = Inches(0.5)
        self._line_width = Pt(2.25)
        self._color = (0, 102, 204)
        self.rounded = False

    @staticmethod
    def _inches_to_float(value: Inches) -> float:
        """Convert Inches to float value in inches."""
        return float(value.inches)

    @staticmethod
    def _pt_to_float(value: Pt) -> float:
        """Convert Pt to float value in points."""
        return float(value.pt)

    @staticmethod
    def _to_inches(value: Union[float, Inches]) -> Inches:
        """Convert a float to Inches if it's not already an Inches object."""
        return Inches(value) if isinstance(value, (int, float)) else value

    @staticmethod
    def _to_pt(value: Union[float, Pt]) -> Pt:
        """Convert a float to Pt if it's not already a Pt object."""
        return Pt(value) if isinstance(value, (int, float)) else value

    @property
    def top_margin(self) -> Inches:
        return self._top_margin

    @top_margin.setter
    def top_margin(self, value: Union[float, Inches]) -> None:
        self._top_margin = self._to_inches(value)

    @property
    def left_margin(self) -> Inches:
        return self._left_margin

    @left_margin.setter
    def left_margin(self, value: Union[float, Inches]) -> None:
        self._left_margin = self._to_inches(value)

    @property
    def right_margin(self) -> Inches:
        return self._right_margin

    @right_margin.setter
    def right_margin(self, value: Union[float, Inches]) -> None:
        self._right_margin = self._to_inches(value)

    @property
    def bottom_margin(self) -> Inches:
        return self._bottom_margin

    @bottom_margin.setter
    def bottom_margin(self, value: Union[float, Inches]) -> None:
        self._bottom_margin = self._to_inches(value)

    @property
    def h_center_margin(self) -> Inches:
        return self._h_center_margin

    @h_center_margin.setter
    def h_center_margin(self, value: Union[float, Inches]) -> None:
        self._h_center_margin = self._to_inches(value)

    @property
    def v_center_margin(self) -> Inches:
        return self._v_center_margin

    @v_center_margin.setter
    def v_center_margin(self, value: Union[float, Inches]) -> None:
        self._v_center_margin = self._to_inches(value)

    @property
    def line_width(self) -> Pt:
        return self._line_width

    @line_width.setter
    def line_width(self, value: Union[float, Pt]) -> None:
        self._line_width = self._to_pt(value)

    @staticmethod
    def color_to_rgb(color_value: HexColorOrName) -> RGBTriplet:
        """Convert a hex color string or a named color to an RGB triplet, ensuring no alpha channel."""
        rgb = ImageColor.getrgb(color_value)
        return rgb[:3]

    @staticmethod
    def validate_rgb(color: RGBTriplet) -> None:
        """Validate that each RGB component is within the 0-255 range."""
        if not all(0 <= c <= 255 for c in color):
            raise ValueError(f"Each color component must be between 0 and 255: {color}")

    @property
    def color(self) -> RGBTriplet:
        """Get the current RGB color."""
        return self._color

    @color.setter
    def color(self, value: Union[RGBTriplet, HexColorOrName]) -> None:
        """Set the color, converting from hex or named color if necessary, and validate it."""
        if isinstance(value, str):
            value = self.color_to_rgb(value)

        value = value[:3]

        self.validate_rgb(value)
        self._color = value

    def __repr__(self) -> str:
        return (
            f"PPTXSettings("
            f"top_margin={self._inches_to_float(self.top_margin):.2f} in, "
            f"left_margin={self._inches_to_float(self.left_margin):.2f} in, "
            f"right_margin={self._inches_to_float(self.right_margin):.2f} in, "
            f"bottom_margin={self._inches_to_float(self.bottom_margin):.2f} in, "
            f"h_center_margin={self._inches_to_float(self.h_center_margin):.2f} in, "
            f"v_center_margin={self._inches_to_float(self.v_center_margin):.2f} in, "
            f"line_width={self._pt_to_float(self.line_width):.2f} pt, "
            f"color={self.color}, "
            f"rounded={self.rounded})"
        )
