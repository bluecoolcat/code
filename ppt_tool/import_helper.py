# Helper script to ensure PyInstaller includes all necessary modules
import moviepy
import moviepy.editor
from moviepy.editor import *
import moviepy.video.io.ffmpeg_reader
import moviepy.audio.io.readers
import moviepy.video.fx.all
import moviepy.audio.fx.all
import numpy
import PIL
from PIL import Image, ImageDraw, ImageFont
import pyttsx3
import pyttsx3.drivers
import pyttsx3.drivers.sapi5
import win32com.client
import websocket
import requests
print('All required modules imported successfully!')
