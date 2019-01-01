#!/usr/bin/env python
# -*- coding: utf-8 -*-

import os
import cv2
import numpy as np
import libs.common as common


class imLib():
    def __init__(self):
        pass
    def set_log(self, log_file):
        common.log_file = log_file

    def find_contrast_image(self, img_fn, templ_fn): # не доделана функция, надо вставить процесс залогинивания (через АСЕ)
        img_fn_c = img_fn[0: -4] + "_contrast.png"
        if not os.path.isfile(img_fn_c):
            main_image = cv2.imread(img_fn, 0)
            main_image = cv2.adaptiveThreshold(main_image, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                            cv2.THRESH_BINARY, 11, 2)
            res = cv2.imwrite(img_fn_c, main_image)
            if not res:
                raise Exception("Error saving image + " + img_fn)
        else:
            main_image = cv2.imread(img_fn_c)

        templ_fn = "c:/AutoClickExtreme/pics/" + templ_fn + ".png"
        templ_fn_c = templ_fn[0: -4] + "_contrast.png"
        if not os.path.isfile(templ_fn_c):
            template = cv2.imread(templ_fn,0)
            template = cv2.adaptiveThreshold(template, 255, cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                                            cv2.THRESH_BINARY, 11, 2)
            res = cv2.imwrite(templ_fn_c, template)
            if not res:
                raise Exception("Error saving image + " + img_fn)
        else:
            template = cv2.imread(templ_fn_c)
        h = template.shape[0]
        w = template.shape[1]

        res = cv2.matchTemplate(main_image, template, cv2.TM_CCOEFF_NORMED)
        threshold = .55
        loc = np.where(res >= threshold)
        pt_arr = []
        for pt in zip(*loc[::-1]):  # Switch collumns and rows
            common.log_print(pt)
            pt_arr.append(pt)
            cv2.rectangle(main_image, pt, (pt[0] + w, pt[1] + h), (0, 0, 255), 2)

        if len(pt_arr):
            cv2.imwrite(img_fn[0:-4] + '_result.png', main_image)
            if len(pt_arr) > 1:
                raise Exception("too big array of matchTemplate result")
        return pt_arr

    def find_image(self, img_fn, templ_fn, threshold): # не доделана функция, надо вставить процесс залогинивания (через АСЕ)
        main_image = cv2.imread(img_fn)
        full_templ_fn = "c:/AutoClickExtreme/pics/" + templ_fn + ".png"
        template = cv2.imread(full_templ_fn)
        if not type(template) == np.ndarray:
            raise Exception("Not found image " + full_templ_fn)
        h = template.shape[0]
        w = template.shape[1]
        res = cv2.matchTemplate(main_image, template, cv2.TM_CCOEFF_NORMED)

        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)

        if max_val > threshold:
            common.log_print("best image match = %f for %s" % (max_val, templ_fn))
            bottom_right = (max_loc[0] + w, max_loc[1] + h)
            cv2.rectangle(main_image, max_loc, bottom_right, (0, 0, 255), 2)
            cv2.imwrite(img_fn[0:-4] + '_result.png', main_image)
            return max_loc
        return False