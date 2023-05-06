import numpy as np
import cv2
from boxdetect import config
from boxdetect.pipelines import get_checkboxes
from operator import itemgetter
import pandas as pd


cfg = config.PipelinesConfig()
cfg.width_range = (20, 55)
cfg.height_range = (20, 40)
cfg.scaling_factors = [0.7675]
cfg.wh_ratio_range = (0.5, 1.7)
cfg.group_size_range = (2, 50)
cfg.dilation_iterations = 0


def extract_boxes(filename):
    img = cv2.imread(filename, 0)
    thresh, img_bin = cv2.threshold(img, 128, 255, cv2.THRESH_BINARY)
    img_bin = 255 - img_bin
    img_bin1 = 255 - img
    thresh1, img_bin1_otsu = cv2.threshold(img_bin1, 128, 255, cv2.THRESH_OTSU)
    img_bin2 = 255 - img
    thresh1, img_bin_otsu = cv2.threshold(img_bin2, 128, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (2, 2))
    vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1, np.array(img).shape[1] // 100))
    eroded_image = cv2.erode(img_bin_otsu, vertical_kernel, iterations=3)
    vertical_lines = cv2.dilate(eroded_image, vertical_kernel, iterations=3)
    hor_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (np.array(img).shape[1] // 120, 1))
    horizontal_lines = cv2.erode(img_bin, hor_kernel, iterations=5)
    horizontal_lines = cv2.dilate(horizontal_lines, hor_kernel, iterations=5)
    vertical_horizontal_lines = cv2.addWeighted(vertical_lines, 0.5, horizontal_lines, 0.5, 0.0)
    vertical_horizontal_lines = cv2.erode(~vertical_horizontal_lines, kernel, iterations=3)
    thresh, vertical_horizontal_lines = cv2.threshold(vertical_horizontal_lines, 128, 255,
                                                      cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    contours, hierarchy = cv2.findContours(vertical_horizontal_lines, cv2.RETR_TREE, cv2.CHAIN_APPROX_SIMPLE)
    boundingBoxes = [cv2.boundingRect(contour) for contour in contours]
    (contours, boundingBoxes) = zip(*sorted(zip(contours, boundingBoxes), key=lambda x: x[1][1]))
    boxes = []
    for contour in contours:
        x, y, w, h = cv2.boundingRect(contour)
        if (w < 1000 and h < 500):
            # image = cv2.rectangle(img,(x,y),(x+w,y+h),(0,255,0),2)
            boxes.append([x, y, w, h])
    boxes = sorted(boxes, key=itemgetter(1))
    return boxes


def extract_checkboxes(file_path):
    checkboxes = get_checkboxes(
        file_path, cfg=cfg, px_threshold=.1, plot=False, verbose=False)
    return checkboxes


def read_docx_table(document,table_num=1,nheader=1):
  table = document.tables[table_num-1]
  data = [[cell.text for cell in row.cells] for row in table.rows]
  df = pd.DataFrame(data)
  if nheader ==1:
    df = df.rename(columns=df.iloc[0]).drop(df.index[0]).reset_index(drop=True)
  if nheader ==2:
    outside_col, inside_col = df.iloc[0], df.iloc[1]
    hier_index = pd.MultiIndex.from_tuples(list(zip(outside_col, inside_col)))
    df = pd.DataFrame(data,columns=hier_index).drop(df.index[[0,1]] ).reset_index(drop=True)
  return df

