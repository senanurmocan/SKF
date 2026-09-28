# -*- coding: utf-8 -*-
"""pytest configuration — sys.path'e proje dizini ekle."""
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
