import mss # fast screen capture
import cv2 # image and video processing
import numpy as np
from paddleocr import PaddleOCR # image text reader


class DeathDetector:
        
    def __init__(self):
        self.sct = mss.mss()
        self.ocr = PaddleOCR(use_angle_cls=False, lang="en", show_log=False)
        self.trigger = "YOU DIED"
        self.primary = self.sct.monitors[2]
        self.already_dead = False # Helps get rid of double counts


    #------------------------#
    #    Runs the Class      #
    #------------------------#
    def detect_death(self):
        """Reutrn True if death detected"""

        image = self.sct.grab(self.get_region())
        return self.check_death(image)

    #------------------------#
    #   Fullscreen -> Box    #
    #------------------------#
    def get_region(self):
        box_width = 800
        box_height = 400

        region = {
            "top": int(self.primary["top"] + (self.primary["height"] - box_height) / 2),
            "left": int(self.primary["left"] + (self.primary["width"] - box_width) / 2),
            "width": box_width,
            "height": box_height,
        }
        return region
    
    #------------------------#
    # Check Image for Death  #
    #------------------------#
    def check_death(self, image):
        """Use OCR to read text"""

        is_dead = False

        # Gets photo "prepped"
        image_np = np.array(image)
        image_bgr = cv2.cvtColor(image_np, cv2.COLOR_BGRA2BGR)

        result = self.ocr.predict(image_bgr, cls=False)
        text_detected_frame = False

        # Run through photo
        if result and result[0]:
            for line in result[0]:
                detected_text = line[1][0]

                if self.trigger.lower() in detected_text.lower():
                    text_detected_frame = True
                    break

        #-----------------------#
        #  State Machine Logic  #
        #-----------------------#

        # Death not yet counted yet
        if text_detected_frame and not self.already_dead:
            self.already_dead = True
            return True
        
        # Reset state
        elif not text_detected_frame and self.already_dead:
            self.already_dead = False
            return False
        
        # Count already counted (Still in death screen) OR no text and alive
        else:
            return False