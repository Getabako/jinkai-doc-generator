"""
Vercel serverless function entry point
"""
import sys
import os

# Add project root to path so we can import tool modules
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from tool.app import app
