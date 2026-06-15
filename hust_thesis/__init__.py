"""
HUST 博士论文生成工具
"""
from .generator import generate_thesis
from .formatters import HustBibFormatter, HustCitationFormatter

__all__ = ["generate_thesis", "HustBibFormatter", "HustCitationFormatter"]
__version__ = "0.1.0"
