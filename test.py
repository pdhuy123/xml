import cv2
import torch
from detectron2.config import get_cfg
from detectron2.engine import DefaultPredictor
from table_transformer import add_table_config
from detectron2.utils.visualizer import Visualizer
import matplotlib.pyplot as plt

# Load config
cfg = get_cfg()
add_table_config(cfg)

cfg.merge_from_file("configs/DLA_table_detection.yaml")
cfg.MODEL.WEIGHTS = "weights/DLA_R_50_FPN_3x.pth"  # tải về từ Microsoft
cfg.MODEL.DEVICE = "cuda" if torch.cuda.is_available() else "cpu"
cfg.MODEL.ROI_HEADS.SCORE_THRESH_TEST = 0.7

predictor = DefaultPredictor(cfg)

# Load image
img_path = "table_image.png"  # ảnh bảng bạn có
im = cv2.imread(img_path)

# Dự đoán bảng
outputs = predictor(im)

# Hiển thị bảng
v = Visualizer(im[:, :, ::-1])
out = v.draw_instance_predictions(outputs["instances"].to("cpu"))
plt.imshow(out.get_image())
plt.axis('off')
plt.show()
