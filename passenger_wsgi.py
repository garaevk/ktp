# -*- coding: utf-8 -*-
"""
WSGI entry point for Beget hosting.
"""
import os
import sys

# Add the application directory to the Python path
sys.path.insert(0, os.path.dirname(__file__))

from app import app as application
