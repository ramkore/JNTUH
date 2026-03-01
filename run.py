#!/usr/bin/env python
"""
Root-level launcher for B.Tech Internal Marks Automation System.
Run this from the project root: python run.py
"""
import sys
import os

# Add src/ to Python path so all imports resolve correctly
src_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'src')
if src_dir not in sys.path:
    sys.path.insert(0, src_dir)

# Change working directory to src/ so asset paths resolve correctly
os.chdir(src_dir)

from main import main
main()
