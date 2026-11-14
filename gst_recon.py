#!/usr/bin/env python3
"""Wrapper to run gst_recon directly from project root."""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(os.path.abspath(__file__)), "src"))
from gst_recon import main

if __name__ == '__main__':
    sys.exit(main())
