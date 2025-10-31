''' © 2023.09.05 kibon-Ku <rlqhs7379@gmail.com> all rights reserved. '''


import sys, os 
import re
import cv2, numpy as np
import pandas as pd
import json
from PIL import Image, ImageDraw

from PyQt5.QtCore import Qt, QPointF
from PyQt5.QtGui import (
    QBrush, 
    QPainter, 
    QPen, 
    QFont, 
    QPixmap, 
    QColor, 
    QScreen, 
    QImage, 
    QKeySequence
)
from PyQt5.QtWidgets import (
    QApplication,
    QGraphicsItem,
    QGraphicsRectItem,
    QGraphicsLineItem,
    QGraphicsScene,
    QGraphicsView,
    QHBoxLayout,
    QPushButton,
    QSlider,
    QVBoxLayout,
    QWidget,
    QFileDialog,
    QLineEdit,
    QTableWidgetItem,
    QTableWidget,
    QMessageBox,
    QSpacerItem,
    QSizePolicy,
    QLabel,
    QShortcut,
    QAbstractItemView,
    QHeaderView,
    QErrorMessage
)


class KUWindow(QWidget):
    
    def __init__(self):
        super().__init__()
        
        '''No exit when there is any exception in PyQt5'''  
        # Back up the reference to the exceptionhook
        sys._excepthook = sys.excepthook
        # Set the exception hook to our wrapping function
        sys.excepthook = self.my_exception_hook
        '''No exit when there is any exception in PyQt5'''
        
        self.setWindowTitle("Grass VI Computation Macro 4.0 (KU)")
        # self.setWindowIcon(QIcon('./icon2.png'))
        
        ''' [Function 4] key button shortcut ''' 
        self.shortcut_create = QShortcut(QKeySequence('Ctrl+C'), self)  # @ https://doc.qt.io/qtforpython-5/PySide2/QtGui/QKeySequence.html 
        self.shortcut_create.activated.connect(self.create_box)
        
        self.shortcut_del = QShortcut(QKeySequence('Del'), self)
        self.shortcut_del.activated.connect(self.delete_box)
        ''' [Function 4] key button shortcut '''
        
        ''' [Function 5] key and mouse wheel shortcut '''
        self.bCtrl = False
        self.deg = QPointF()
        self.zoom = QPointF()
        ''' [Function 5] key and mouse wheel shortcut '''
        
        ### Initialization
        # layout size
        self.w, self.h = 1150, 880                 
        # box
        self.box = None   
        self.box_id = 0  # box.setData(self.box_id, self.box_id)  # void QGraphicsItem::setData(int key, const QVariant &value)  @ https://doc.qt.io/qt-5/qgraphicsitem.html#setData
        self.box_dict = {}  # box coordination   
        self.box_vi_dict = {}
        self.box_json_dict = {}
        self.box_vi_list = []
        self.box_json_list = []
        # name
        self.folder_name = None
        self.file_name = None
        self.save_dir = None
        self.data = []
        self.json = []
        # path
        self.folder_path = None
        self.file_path = None
        # cell info
        self.cell_row = 0
        self.cell_col = 0
        
        ## set widget size
        self.setGeometry(0,0,1900,1000)
        self.setFixedSize(1900,1000)  # fixed the size of the widget

        # Defining a scene rect of wxh, with it's origin at 0,0.
        # If we don't set this on creation, we can set it later with .setSceneRect
        self.scene = QGraphicsScene(0, 0, self.w, self.h)
        self.view = QGraphicsView(self.scene)
        self.view.setRenderHint(QPainter.Antialiasing)  # https://wikidocs.net/36648        
        
        ## Define layout.
        parbox = QVBoxLayout()
        self.setLayout(parbox)
                
        vbox = QVBoxLayout()  # vertical  @ https://doc.qt.io/qt-5/qvboxlayout.html
        vbox2 = QVBoxLayout()  # vertical  @ https://doc.qt.io/qt-5/qvboxlayout.html
        vbox.setAlignment(Qt.AlignTop)
        
        hbox1 = QHBoxLayout()  # horizontal 
        hbox2 = QHBoxLayout()  
        hbox3 = QHBoxLayout()  
        hbox4 = QHBoxLayout()  
        hbox5 = QHBoxLayout()  
        hbox6 = QHBoxLayout()  
        
        #### Define Button 
        ## gbox ; grid
        folder = QPushButton("Input Folder")  ## db folder selection
        folder.setStyleSheet("color: black;"
                        "background-color: cyan")
        folder.clicked.connect(self.select_folder)        
        self.folder_text = QLineEdit()  ## db folder selection
        
        save_folder = QPushButton("Output Folder")
        save_folder.setStyleSheet("color: black;"
                        "background-color: cyan")
        save_folder.clicked.connect(self.select_save_folder)
        self.save_folder_text = QLineEdit()  ## db folder selection
        
        rgb_info = QPushButton("Save RGB and RoI")
        rgb_info.setStyleSheet("color: black;"
                        "background-color: cyan")
        rgb_info.clicked.connect(self.save_rgb_roi)
               
        vi_info = QPushButton("Save VI and RoI")
        vi_info.setStyleSheet("color: black;"
                        "background-color: cyan")
        vi_info.clicked.connect(self.save_vi_roi)
                    
        exl = QPushButton("Export DB")
        exl.setStyleSheet("color: black;"
                        "background-color: cyan")
        exl.clicked.connect(self.save_excel)
        
        # filelist table
        self.filelist = QTableWidget(0,1)
        self.filelist.setHorizontalHeaderLabels(['FileName']) 
        self.filelist.setEditTriggers(QAbstractItemView.NoEditTriggers) 
        self.filelist.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.filelist.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.filelist.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  
        self.filelist.cellClicked.connect(self.select_cell)
        
        spaceItem1 = QSpacerItem(100, 0, QSizePolicy.Expanding)   
        
        hbox1.addWidget(folder, 2) 
        hbox1.addWidget(self.folder_text, 6)
        hbox1.addWidget(save_folder, 2)
        hbox1.addWidget(self.save_folder_text, 6)
        hbox1.addSpacerItem(spaceItem1)
        hbox1.addWidget(rgb_info, 2)
        hbox1.addWidget(vi_info, 2)
        hbox1.addWidget(exl, 2)
        
        hbox1.setSpacing(10)
        parbox.addLayout(hbox1)

        # hbox3 ; horizontal
        ''' [Function 1] box size update '''
        w_box, h_box = 65, 65  # default size # replace
        
        self.label_w_box = QLabel(self)
        self.label_w_box.setText('Box Width : ')
        self.line_edit_box_w = QLineEdit(str(w_box), self)
        
        self.label_h_box = QLabel(self)
        self.label_h_box.setText('Box Height : ')
        self.line_edit_box_h = QLineEdit(str(h_box), self)        
        ''' [Function 1] box size update '''

        ''' [Function 2] box rep update '''
        self.label_box_rep = QLabel(self)
        self.label_box_rep.setText('Box REP : ')
        self.line_edit_box_rep = QLineEdit('1', self)   # replace # default is just a basic box
        ''' [Function 2] box rep update '''
                
        ''' [Function 3] box name update '''
        self.label_box_id = QLabel(self)
        self.label_box_id.setText('Box Name : ')
        self.line_edit_box_id = QLineEdit(self)      
        ''' [Function 3] box name update '''
        
        hbox3.addWidget(self.label_w_box, 1) 
        hbox3.addWidget(self.line_edit_box_w, 1)
        hbox4.addWidget(self.label_h_box, 1)
        hbox4.addWidget(self.line_edit_box_h, 1)
        hbox5.addWidget(self.label_box_rep, 1)
        hbox5.addWidget(self.line_edit_box_rep, 1)
        hbox6.addWidget(self.label_box_id, 1)
        hbox6.addWidget(self.line_edit_box_id, 1)
                
        ## vbox ; vertical                    
        box = QPushButton("Create (Ctrl+C)")
        box.clicked.connect(self.create_box)
        
        dlt = QPushButton("Delete (Del)")
        dlt.clicked.connect(self.delete_box)
    
        rot = QPushButton("Rotate\n(Ctrl+MouseWheel)")
        
        # slider ; rotation
        rotate = QSlider()
        angle_min, angle_max = 0, 180
        rotate.setRange(angle_min, angle_max)
        rotate.setTickPosition(QSlider.TicksRight)
        rotate.setTickInterval(1)
        rotate.setSingleStep(1)  
        rotate.valueChanged.connect(self.rotate_box)
                
        vbox.addLayout(hbox3)
        vbox.addLayout(hbox4)
        vbox.addLayout(hbox5)
        vbox.addLayout(hbox6)
        vbox.addWidget(box)
        vbox.addWidget(dlt)
        vbox.addWidget(rot)
        vbox.addWidget(rotate)
        
        vbox.setSpacing(10)
        vbox.setContentsMargins(5,0,0,400)
        
        ## hbox ; horiziontal
        spaceItem2 = QSpacerItem(30, 0, QSizePolicy.Expanding)  # QSpacerItem : https://doc.qt.io/qtforpython-5/PySide2/QtWidgets/QSpacerItem.html
        
        hbox2.addWidget(self.filelist, 15)  # ratio
        hbox2.addSpacerItem(spaceItem2)
        hbox2.addLayout(vbox, 10)
        hbox2.addWidget(self.view, 75)
        
        # Show DB Table        
        self.table_db = QTableWidget()
        self.table_db.setColumnCount(1)
        self.table_db.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_db.setHorizontalHeaderLabels(['DB table'])
        self.table_db.setEditTriggers(QAbstractItemView.NoEditTriggers)
        hbox2.addWidget(self.table_db, 15)
        
        hbox2.setSpacing(10)
        hbox2.setContentsMargins(0,20,0,0)
        
        parbox.addLayout(hbox2)        
        parbox.setContentsMargins(30,30,30,30)

    ''' [Function 5] key and mouse wheel shortcuts ''' 
    def keyPressEvent(self, e):
        if e.key() == Qt.Key_Control:
            self.bCtrl = True
        self.update()
 
    def keyReleaseEvent(self, e):
        if e.key() == Qt.Key_Control:
            self.bCtrl = False
        self.update()
        
    def wheelEvent(self, e):
        if self.bCtrl: # rotate box
            self.zoom += e.angleDelta() / 120
            zoomtxt = f'Zoom X:{self.zoom.x()}, Zoom Y:{self.zoom.y()}'
            # print('zoomtxt : ', zoomtxt)
            
            # value = abs(int(self.zoom.y()))
            value = int(self.zoom.y())
            self.rotate_box(value)
        else: 
            self.deg += e.angleDelta() / 8
            degtxt  = f'Deg X:{self.deg.x()}, Deg Y:{self.deg.y()}'
            # print('degtxt : ', degtxt)
 
        self.update()
    ''' [Function 5] key and mouse wheel shortcuts '''
        
    def select_save_folder(self):
        
        self.save_dir = QFileDialog.getExistingDirectory(self, 'Open Folder', 'C:\\')  
        self.save_folder_text .setText(self.save_dir)        
                
    def select_folder(self):
        
        """ select working directory. """
        
        # Save the current folder
        self.folder_path = QFileDialog.getExistingDirectory(self, 'Open Folder', 'C:\\') 
        self.folder_text.setText(self.folder_path)  
        # Save the selected folder name
        self.folder_name = self.folder_path.split('/')[-1]
        
        # show filelist
        FileList = [ x for x in os.listdir(self.folder_path) if re.match('.*\.JPG|.*\.jpg|.*\.PNG|.*\.png|.*\.TIF|.*\.TIFF|.*\.tif|.*\.tiff', x)] 
        FileList.sort()
        
        ## qtablewidget에 출력하기 
        cell_row = len(FileList)
        col = 0
        self.filelist.setRowCount(cell_row)  # https://doc.qt.io/qt-5/qtablewidget.html
        
        for i in range(cell_row):
            FileName = str(FileList[i])  
            
            item = QTableWidgetItem(FileName)
            self.filelist.setItem(i, col, item)

    def select_cell(self): 
        
        #### Initialization
        self.box_dict.clear()
        self.box_0_dict = {}
        self.box_vi_dict = {}
        self.box_id = 0  
        self.scene.clear()  # https://doc.qt.io/archives/qt-4.8/qgraphicsscene.html#clear
        
        cell_Key = self.filelist.selectedIndexes()  # input the selected list

        ## Call the numbers
        self.cell_row = cell_Key[0].row()  # call the selected row
        self.cell_col = cell_Key[0].column()  # call the selcted column

        ## Take the value out from the cell
        self.file_name = self.filelist.item(self.cell_row, self.cell_col).text()  # make the text (string) as variable
        self.file_path = self.folder_path + '/' + self.file_name  # image file name (absolute path)   

        ## call the function displayImg
        self.display_img() 
        
    def display_img(self):
        
        ## Korean path > NoneType Error 
        n = np.fromfile(self.file_path, np.uint8)   # read a binary data as a numpy array
        im = cv2.imdecode(n, cv2.IMREAD_UNCHANGED)  # decode to enable us to process the image using opencv
        
        if im.dtype !=np.uint8:  # vi
            ## convert float32 to uint8
            im[im<0]=0  # normalization to get meaningful data       
            im *= 255  #@times
            im = im.astype(np.uint8)
            
            ## set QPixmap 
            height, width = im.shape
            # print('height, width :', height, width )
            qImg = QImage(im.data, width, height, width, QImage.Format_Grayscale8)
            pixmap = QPixmap.fromImage(qImg)
            pixmap = QPixmap(pixmap)
        else:  # rgb
            pixmap = QPixmap(self.file_path)
            
        pixmap = pixmap.scaled(self.w, self.h)
        pixmapitem = self.scene.addPixmap(pixmap)
        pixmapitem.setPos(0, 0)  # its position is initialized to (0, 0).
        
        ''' [Function 5] pixmap size fixing '''
        ## resize image to fit in graphics view : https://forum.qt.io/topic/92586/resize-image-to-fit-in-graphics-view
        self.view.fitInView(pixmapitem)
        
        ''' [Function 5] pixmap size fixing '''   
               
    def create_box(self):
        
        ''' [Function 1] box size update '''
        self.w_box = int(self.line_edit_box_w.text())
        self.h_box = int(self.line_edit_box_h.text())
        ''' [Function 1] box size update '''
        
        ### box item 생성!
        self.box = QGraphicsRectItem(0, 0, self.w_box, self.h_box)  
        # print('area : ', self.box.sceneBoundingRect().topLeft(), self.box.sceneBoundingRect().bottomRight())

        ''' [Function 2] box rep update '''
        self.box_rep_num = int(self.line_edit_box_rep.text())
        self.draw_box_line(self.box_rep_num)
        ''' [Function 2] box rep update '''        
               
        # define the brush
        brush = QBrush(Qt.cyan, Qt.NoBrush)   # Qt.BrushStyle : https://doc.qt.io/archives/qtjambi-4.5.2_01/com/trolltech/qt/core/Qt.BrushStyle.html | https://wikidocs.net/37095
        self.box.setBrush(brush)
        
        # Draw a rectangle item, setting the dimensions.
        self.scene.addItem(self.box)
        
        # Set the self.box item as moveable and selectable.  
        self.box.setFlag(QGraphicsItem.ItemIsMovable)
        self.box.setFlag(QGraphicsItem.ItemIsSelectable)        
        
        ## box_id
        ''' [Function 3] box name update '''
        self.box_id += 1
        self.box_name = self.line_edit_box_id.text() + str(self.box_id)
        ''' [Function 3] box name update '''     
        text = self.scene.addText('ID ' + str(self.box_name), QFont("Times", 11, QFont.Bold))  # Qfont 
        text.setPos(self.box.pos().x(), self.box.pos().y()-20)
        text.setDefaultTextColor(QColor(Qt.cyan))   #html color chart : https://wepplication.github.io/tools/colorPicker/ 
        text.setParentItem(self.box)  # Group box and box_name  
                
        ## add box to box_0_dict
        degree = 0
        self.box_dict[self.box_name] = [degree, self.box]

    ''' [Function 2] box rep update '''
    def draw_box_line(self, N:int):
        
        # print('N : ', N)
        
        # define the pen
        pen = QPen(Qt.cyan)
        pen.setWidth(2)
        self.box.setPen(pen)
        pen.setWidth(2)
        
        ## generate a line item to divide it into N
        for i in range(N):
            line = QGraphicsLineItem(self.w_box/N*(N-i),0, self.w_box/N*(N-i),self.h_box)
            line.setParentItem(self.box)
            line.setPen(pen)
        
    def delete_box(self):
        
        items = self.scene.selectedItems()
        for item in items:
            
            self.scene.removeItem(item) 
            
            # undo
            self.box_id -= 1  
    
    def RotateCoordinate(self, org_pt, theta, standart_pt):
        import math
    
        rot_x = (org_pt[0] - standart_pt[0])*math.cos(math.radians(theta)) - (org_pt[1] - standart_pt[1])*math.sin(math.radians(theta)) + standart_pt[0]

        rot_y = (org_pt[0] - standart_pt[0])*math.sin(math.radians(theta)) + (org_pt[1] - standart_pt[1])*math.cos(math.radians(theta)) + standart_pt[1]

        return rot_x, rot_y
    
    def rotate_box(self, value):
                
        items = self.scene.selectedItems()
        for item in items:
            
            ## box
            item.setRotation(value)

            box_dict = self.box_dict[self.box_id]
            self.box_dict[self.box_id] = [value, box_dict[1]]
            # print('box dict when rotated : ', self.box_dict[self.box_id][1].pos())
            # print('total box dict : ', self.box_dict)
    
    '''[Function 2] box rep update '''
    def lines_ptrs(self, N:int, x, y, value):
        
        ## lines
        rot_ptrs_list = []
        for i in range(N):
            
            x1, y1 = x+self.w_box/N*(N-i), y
            x2, y2 = x+self.w_box/N*(N-i), y+self.h_box
            
            rot_x1, rot_y1 = self.RotateCoordinate([x1,y1], value, [x,y])
            rot_x2, rot_y2 = self.RotateCoordinate([x2,y2], value, [x,y])  
            
            rot_ptrs_list.insert(0, [[rot_x1, rot_y1], [rot_x2, rot_y2]])
            
        return rot_ptrs_list


    def crop_image_polygon(self, input_image_path, output_image_path, polygon_coords):
        
        original_image = Image.open(input_image_path)
        original_image = original_image.resize((self.w, self.h))
        

        # 새로운 이미지를 만들고 투명한 배경으로 설정
        new_image = Image.new("RGBA", original_image.size, (0, 0, 0, 0))

        # 다각형 좌표로 마스크 생성
        mask = ImageDraw.Draw(new_image)
        mask.polygon(polygon_coords, fill=(255, 255, 255, 255))

        # 다각형 모양으로 이미지 자르기
        result = Image.alpha_composite(original_image.convert("RGBA"), new_image)

        # 결과 이미지 저장
        result.save(output_image_path, "PNG")


    ## RGB 범위만 가능!  # Result = average pixcel value and xy-coordinates                
    def save_rgb_roi(self):
                
        im_rgb = cv2.imread(self.file_path, 1)
        im = cv2.cvtColor(im_rgb, cv2.COLOR_BGR2GRAY)
        
        im_rgb = cv2.resize(im_rgb, (self.w, self.h), cv2.INTER_LINEAR)
        im = cv2.resize(im, (self.w, self.h), cv2.INTER_LINEAR)    
                   
        # bin_total = np.zeros(im.shape, np.uint8) 
        # print('bin_total:', bin_total)
        
        ## Reset the dictionary
        self.box_vi_dict = {} 
        self.box_json_dict = {}  # [sol]
        self.box_json_dict['filename'] = self.file_name  # instead, from collections import OrderedDict
        
        for box_id, box_lines in self.box_dict.items():  # for key,value in dict.items()
            
            value = box_lines[0]
            box = box_lines[1]
            
            x1,y1 = box.pos().x(), box.pos().y()  #  @ https://doc.qt.io/qt-5/qrectf.html
            # x5,y5 = box.pos().x()+self.w_box, box.pos().y()
            # x6,y6 = box.pos().x()+self.w_box, box.pos().y()+self.h_box
            x10,y10 = box.pos().x(), box.pos().y()+self.h_box
                        
            rot_x1, rot_y1 = self.RotateCoordinate([x1,y1], value, [x1,y1])    
            # rot_x5, rot_y5 = self.RotateCoordinate([x5,y5], value, [x1,y1])    
            # rot_x6, rot_y6 = self.RotateCoordinate([x6,y6], value, [x1,y1])    
            rot_x10, rot_y10 = self.RotateCoordinate([x10,y10], value, [x1,y1])    
                        
            ## lines
            rot_ptrs_list = self.lines_ptrs(self.box_rep_num, x1, y1, value)
            rot_ptrs_list.insert(0, [[rot_x1, rot_y1], [rot_x10, rot_y10]])
            
            box_ptrs_list = []
            for i in range(self.box_rep_num):
                
                box_ptrs_list.append(np.array([rot_ptrs_list[i][0], rot_ptrs_list[i+1][0], rot_ptrs_list[i+1][1], rot_ptrs_list[i][1]], np.int32))
                
            # print('box_ptrs_list : ', box_ptrs_list)     
                            
            vi_list = []
            bin_total = np.zeros(im.shape, np.uint8)
            for box_ptrs in box_ptrs_list:
                # print(box_ptrs)
                
                bin = np.zeros(im.shape, np.uint8)
               
                mask = cv2.fillPoly(bin, [box_ptrs], 255, cv2.LINE_AA)
                mask_total = cv2.fillPoly(bin_total, [box_ptrs], 255, cv2.LINE_AA)
                
                # Verify
                # print('im_rgb:', im_rgb.dtype, im_rgb.shape)
                # print('mask_total:', mask_total.dtype, mask_total.shape)
                
                im_box = cv2.bitwise_and(im,im, mask=mask)
                im_box_total = cv2.bitwise_and(im_rgb,im_rgb, mask=mask_total)          

                # sum(vi)/sum(px) 구하기 
                vi_px = cv2.countNonZero(im_box)
                vi_list.append(vi_px)      
                
            # Crop each image: https://www.life2coding.com/cropping-polygon-or-non-rectangular-region-from-image-using-opencv-python/ 
            save_each_path = self.save_dir + '/' + str(box_id) + '_' + self.file_name
            # Flatten the nested list
            flattened_list = []
            for sublist in box_ptrs_list:
                for point in sublist:
                    flattened_list.append(point)

            # Convert the list to a NumPy array if needed
            flattened_array = np.array(flattened_list)
            # Bounding Rectangle
            rect = cv2.boundingRect(flattened_array) # returns (x,y,w,h) of the rect
            # Save the Cropping
            cropped = im_box_total[rect[1]: rect[1] + rect[3], rect[0]: rect[0] + rect[2]]
            cv2.imwrite(save_each_path, cropped)
            
                
            self.box_vi_dict[box_id]= vi_list
            
            conv_box_ptrs_list = [arr.tolist() for arr in box_ptrs_list]
            self.box_json_dict[box_id]= conv_box_ptrs_list
        # print('self.box_vi_dict:', self.box_vi_dict)
        # print('self.box_json_dict:', self.box_json_dict)
       
        ## show on table_db
        self.setTableWidgetData(self.box_vi_dict)
        
        self.box_vi_dict['filename'] = self.file_name
        # self.box_json_dict['filename'] = self.file_name
        
        self.box_vi_list.append(self.box_vi_dict) 
        self.box_json_list.append(self.box_json_dict) 
               
        self.data = self.box_vi_list
        self.json = self.box_json_list
        # print('self.data : ', self.data)
        # print('self.json : ', self.json)
        
        ## 스크린 샷 -> 번호 적힌 이미지 저장
        self.save_img()    
        
        ## 완료한 이미지, 배경색 바꾸기!
        self.cell_bg_color()   
        
        # QMessageBox.information(self, "QMessageBox", "Successfully Saved the Data!")
            
    ## vi 범위만 가능!  # Result = average vi value                
    def save_vi_roi(self):
                
        im = cv2.imread(self.file_path, cv2.IMREAD_UNCHANGED)
        im = cv2.resize(im, (self.w, self.h), cv2.INTER_LINEAR)  
                   
        bin_total = np.zeros(im.shape, np.uint8) 
        
        ## Reset the dictionary
        self.box_vi_dict = {} 
        self.box_json_dict = {}  # [sol]
        self.box_json_dict['filename'] = self.file_name  # instead, from collections import OrderedDict
        
        for box_id, box_lines in self.box_dict.items():  # for key,value in dict.items()
            
            value = box_lines[0]
            box = box_lines[1]
            
            x1,y1 = box.pos().x(), box.pos().y()  #  @ https://doc.qt.io/qt-5/qrectf.html
            # x5,y5 = box.pos().x()+self.w_box, box.pos().y()
            # x6,y6 = box.pos().x()+self.w_box, box.pos().y()+self.h_box
            x10,y10 = box.pos().x(), box.pos().y()+self.h_box
                        
            rot_x1, rot_y1 = self.RotateCoordinate([x1,y1], value, [x1,y1])    
            # rot_x5, rot_y5 = self.RotateCoordinate([x5,y5], value, [x1,y1])    
            # rot_x6, rot_y6 = self.RotateCoordinate([x6,y6], value, [x1,y1])    
            rot_x10, rot_y10 = self.RotateCoordinate([x10,y10], value, [x1,y1])    
                        
            ## lines
            rot_ptrs_list = self.lines_ptrs(self.box_rep_num, x1, y1, value)
            rot_ptrs_list.insert(0, [[rot_x1, rot_y1], [rot_x10, rot_y10]])
            
            box_ptrs_list = []
            for i in range(self.box_rep_num):
                
                box_ptrs_list.append(np.array([rot_ptrs_list[i][0], rot_ptrs_list[i+1][0], rot_ptrs_list[i+1][1], rot_ptrs_list[i][1]], np.int32))
                
            # print('box_ptrs_list : ', box_ptrs_list)     
            
            im[im<0] = 0   # min 0 
            im[im>1] = 1   # max 1 
                            
            vi_list = []
            for box_ptrs in box_ptrs_list:
                # print(box_ptrs)
                
                bin = np.zeros(im.shape, np.uint8)
                mask = cv2.fillPoly(bin, [box_ptrs], 255, cv2.LINE_AA)
                mask_total = cv2.fillPoly(bin_total, [box_ptrs], 255, cv2.LINE_AA)
                im_box = cv2.bitwise_and(im,im, mask=mask)
                im_box_total = cv2.bitwise_and(im,im, mask=mask_total)          

                # sum(vi)/sum(px) 구하기 
                vi_px = cv2.countNonZero(im_box)
                vi_sum1 = sum(im_box)              
                vi_sum = sum(vi_sum1)
                vi = vi_sum/vi_px
                vi_list.append(vi)      
                
            self.box_vi_dict[box_id]= vi_list
            
            conv_box_ptrs_list = [arr.tolist() for arr in box_ptrs_list]
            self.box_json_dict[box_id]= conv_box_ptrs_list
        # print('self.box_vi_dict:', self.box_vi_dict)
        # print('self.box_json_dict:', self.box_json_dict)
       
        ## show on table_db
        self.setTableWidgetData(self.box_vi_dict)
        
        self.box_vi_dict['filename'] = self.file_name
        # self.box_json_dict['filename'] = self.file_name
        
        self.box_vi_list.append(self.box_vi_dict) 
        self.box_json_list.append(self.box_json_dict) 
               
        self.data = self.box_vi_list
        self.json = self.box_json_list
        # print('self.data : ', self.data)
        # print('self.json : ', self.json)
        
        ## 스크린 샷 -> 번호 적힌 이미지 저장
        self.save_img()    
        
        ## 완료한 이미지, 배경색 바꾸기!
        self.cell_bg_color()   
        
        # QMessageBox.information(self, "QMessageBox", "Successfully Saved the Data!") 
        
    ## show on table_db
    def setTableWidgetData(self, dict):
        
        column_headers = ['ID', 'AVG VI']
        self.table_db.setRowCount(100)
        self.table_db.setColumnCount(2)
        # self.table_db.setColumnWidth(0, 115)
        self.table_db.setHorizontalHeaderLabels(column_headers)
        self.table_db.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        # Reset past data
        self.table_db.clearContents()

        row = 0
        for k, v in dict.items():
            for _, val in enumerate(v):
                
                row += 1
                
                k = str(k)
                val = str(np.round(val, 5))  
                # print('row, val : ', row, val)
                
                item_id = QTableWidgetItem(k)
                item_vi = QTableWidgetItem(val)

                self.table_db.setItem(row, 0, item_id)
                self.table_db.setItem(row, 1, item_vi)
            
    def save_img(self):
        
        try:
            # path
            save_img_path = os.path.join(self.save_dir, self.file_name)
            
            # ScreenShot 
            p = QScreen.grabWindow(app.primaryScreen(), w.winId())  #(main, current)
            p.save(save_img_path, 'jpg')
            
        except TypeError as e:
            QMessageBox.information(self, "Error", 'Make the Output Folder Path!')

        # crop img save
        # cv2.imwrite(save_mask_path, im_box_total)
                
    def cell_bg_color(self):
    
        self.filelist.item(self.cell_row, self.cell_col).setBackground(Qt.cyan)
    
    def save_excel(self):
        
        df = pd.DataFrame(self.data)
        df.set_index('filename', inplace=True)   # filename 열을 인덱스로 지정
        
        # Save path
        save_exl_path = os.path.join(self.save_dir, self.folder_name + '.xlsx') 
        save_json_path = os.path.join(self.save_dir, self.folder_name + '.json') 
        
        ## Save VI as Excel file
        writer = pd.ExcelWriter(save_exl_path, engine='xlsxwriter')
        df.to_excel(writer, na_rep = 'NaN', sheet_name=self.folder_name )
        writer.save()
        
        ## Save 
        df_json = self.json
        with open(save_json_path, "w") as json_file:
            json.dump(df_json, json_file)
            
        QMessageBox.information(self, "QMessageBox", "Successfully Saved All Datas as Excel!")
    
    def my_exception_hook(exctype, value, traceback):
        
        '''No exit when there is any exception in PyQt5'''

        # show the error message box to notify to users
        if exctype == TypeError:
            # show the error message box
            error_message = "A TypeError occurred: " + str(value)

        else:   # call the default exception hook
            # Print the error and traceback
            # Call the normal Exception hook after
            sys._excepthook(exctype, value, traceback)


if __name__ == "__main__":

    app = QApplication(sys.argv)
    w = KUWindow()
    w.show()
    sys.exit(app.exec())
 
