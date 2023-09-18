''' © 2023.09.05 kibon-Ku <rlqhs7379@gmail.com> all rights reserved. '''


import sys, os 
import re
import cv2, numpy as np
import pandas as pd
import json

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
        
        '''pyqt5 에서 exception 발생시 종료 방지 방법'''  # @ https://4uwingnet.tistory.com/13
        # Back up the reference to the exceptionhook
        sys._excepthook = sys.excepthook
        # Set the exception hook to our wrapping function
        sys.excepthook = self.my_exception_hook
        '''pyqt5 에서 exception 발생시 종료 방지 방법'''
        
        self.setWindowTitle("Grass VI Computation Macro 4.0 (KU)")
        # self.setWindowIcon(QIcon('./icon2.png'))
        
        ''' [추가 기능 4] key button shortcut '''  # @ https://learndataanalysis.org/create-and-assign-keyboard-shortcuts-to-your-pyqt-application-pyqt5-tutorial/ 
        self.shortcut_create = QShortcut(QKeySequence('Ctrl+C'), self)  # @ https://doc.qt.io/qtforpython-5/PySide2/QtGui/QKeySequence.html 
        self.shortcut_create.activated.connect(self.create_box)
        
        self.shortcut_del = QShortcut(QKeySequence('Del'), self)
        self.shortcut_del.activated.connect(self.delete_box)
        ''' [추가 기능 4] key button shortcut '''
        
        ''' [추가 기능 5] key and mouse wheel shortcut '''
        self.bCtrl = False
        self.deg = QPointF()
        self.zoom = QPointF()
        ''' [추가 기능 5] key and mouse wheel shortcut '''
        
        ### 변수 선언
        # layout size
        self.w, self.h = 1150, 880   # img size..?                  
        # box
        self.box = None   
        self.box_id = 0  # box.setData(self.box_id, self.box_id)  # void QGraphicsItem::setData(int key, const QVariant &value)  @ https://stackoverflow.com/questions/70353754/how-to-identify-a-qgraphicsitem | https://doc.qt.io/qt-5/qgraphicsitem.html#setData
        self.box_dict = {}  # box coordination   
        # self.box_ptrs = {}  # 
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
        self.setFixedSize(1900,1000)  # 창 크기 고정. 동적 x

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
        # self.filelist.setColumnWidth(0, 230)  # ah! start from '0'  @ https://stackoverflow.com/questions/3172026/qt-setcolumnwidth-does-not-work
        self.filelist.setHorizontalHeaderLabels(['FileName']) 
        self.filelist.setEditTriggers(QAbstractItemView.NoEditTriggers)  # edit 금지 모드
        self.filelist.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        self.filelist.setHorizontalScrollBarPolicy(Qt.ScrollBarAsNeeded)
        # self.filelist.setVerticalScrollBarPolicy(Qt.ScrollBarAlwaysOn)
        self.filelist.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)  # 테이블 크기에 맞게 자동 스트레치
        # self.filelist.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeToContents)  # 내용 길이에 맞게 리사이즈
        self.filelist.cellClicked.connect(self.select_cell)
        
        spaceItem1 = QSpacerItem(100, 0, QSizePolicy.Expanding)   # @ https://www.programcreek.com/python/example/108080/PyQt5.QtWidgets.QSpacerItem
        
        hbox1.addWidget(folder, 2)  # https://lifeiseggs.tistory.com/866
        hbox1.addWidget(self.folder_text, 6)
        hbox1.addWidget(save_folder, 2)
        hbox1.addWidget(self.save_folder_text, 6)
        hbox1.addSpacerItem(spaceItem1)
        hbox1.addWidget(vi_info, 2)
        hbox1.addWidget(exl, 2)
        
        hbox1.setSpacing(10)
        parbox.addLayout(hbox1)

        # hbox3 ; horizontal
        ''' [추가 기능 1] box size update '''
        w_box, h_box = 185, 125   # default size > 잔디 기준
        
        self.label_w_box = QLabel(self)
        self.label_w_box.setText('Box Width : ')
        self.line_edit_box_w = QLineEdit(str(w_box), self)
        
        self.label_h_box = QLabel(self)
        self.label_h_box.setText('Box Height : ')
        self.line_edit_box_h = QLineEdit(str(h_box), self)        
        ''' [추가 기능 1] box size update '''

        ''' [추가 기능 2] box rep update '''
        self.label_box_rep = QLabel(self)
        self.label_box_rep.setText('Box REP : ')
        self.line_edit_box_rep = QLineEdit('4', self)   # default is just a basic box
        ''' [추가 기능 2] box rep update '''
                
        ''' [추가 기능 3] box name update '''
        self.label_box_id = QLabel(self)
        self.label_box_id.setText('Box Name : ')
        self.line_edit_box_id = QLineEdit(self)      
        ''' [추가 기능 3] box name update '''
        
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
        rotate.setSingleStep(1)  # https://stackoverflow.com/questions/4827885/qslider-stepping 
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
        
        hbox2.addWidget(self.filelist, 15)  # 비율 ; https://stackoverflow.com/questions/42367689/pyqt-unequally-divide-the-area-occupied-by-widgets-in-a-qhboxlayout
        hbox2.addSpacerItem(spaceItem2)
        hbox2.addLayout(vbox, 10)
        hbox2.addWidget(self.view, 75)
        
        # Show DB Table        
        self.table_db = QTableWidget()
        self.table_db.setColumnCount(1)
        self.table_db.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table_db.setHorizontalHeaderLabels(['DB table'])
        self.table_db.setEditTriggers(QAbstractItemView.NoEditTriggers)  # edit 금지 모드
        hbox2.addWidget(self.table_db, 15)
        
        hbox2.setSpacing(10)
        hbox2.setContentsMargins(0,20,0,0)
        
        '''
        ## Explanation : https://wikidocs.net/33889
        lbl_explain = QLabel()
        lbl_explain_title = QLabel()
        lbl_explain_title.setStyleSheet("color: blue;"
                              "font-weight: bold")
        lbl_explain.setStyleSheet("color: black;"
                              "background-color: cyan;"
                              "border-style: dashed;"
                              "border-width: 3px;"
                              "border-color: #1E90FF")
        lbl_explain_title.setText("** HOW TO **")
        lbl_explain.setText( "\n" + \
            "1. Click the <Input Folder> button and Selcet the Folder what you want to process.\n" + \
            "2. Click the <Output Folder> button and Select the Folder what you want to save the data in.\n" + \
            "3. Click a cell of filelist below the <FileName> what you want to show an Image.\n" + \
            "4. Click the <Create> button when you want to make a box\n" + \
            "5. Click the <Delete> button when you want to remove a box you selected.\n" + \
            "6. Click the <Rotate> button when you want to rotate a box you selected.\n" + \
            "7. Click the <Save X Info> button when you finish processing the image. X is RGB or vi. It must do one by one.\n" + \
            "8. Repeate the steps of '3.~7.' about all images.\n" + \
            "9. Click the <Save them as Excel> button when you finish processing all images. You can use it if you want to check the data you have done.\n")
        vbox2.addWidget(lbl_explain_title)
        vbox2.addWidget(lbl_explain)
        vbox2.setSpacing(10)
        vbox2.setContentsMargins(0,50,0,0)
        '''
        
        parbox.addLayout(hbox2)        
        parbox.setContentsMargins(30,30,30,30)

    ''' [추가 기능 5] key and mouse wheel shortcuts ''' 
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
        else: # 이건 다른 기능으로 쓰게 남겨두기
            self.deg += e.angleDelta() / 8
            degtxt  = f'Deg X:{self.deg.x()}, Deg Y:{self.deg.y()}'
            # print('degtxt : ', degtxt)
 
        self.update()
    ''' [추가 기능 5] key and mouse wheel shortcuts '''
        
    def select_save_folder(self):
        
        self.save_dir = QFileDialog.getExistingDirectory(self, 'Open Folder', 'C:\\')  # PyQT UIfh 디렉토리/파일 경로 얻기  # http://jiwonkun.blogspot.com/2013/12/pyqt-ui.html
        self.save_folder_text .setText(self.save_dir)  # 폴더 경로, <edit_Folder>에 출력하기       
                
    def select_folder(self):
        
        """ select working directory. """
        
        # 폴더 경로 저장 > 개발 중에는, 고정된 주소로
        self.folder_path = QFileDialog.getExistingDirectory(self, 'Open Folder', 'C:\\')  # PyQT UIfh 디렉토리/파일 경로 얻기  # http://jiwonkun.blogspot.com/2013/12/pyqt-ui.html
        # self.folder_path = 'D:/4.Turf_re/3_3-6_5_nir+otsu'
        self.folder_text.setText(self.folder_path)  # 폴더 경로, <edit_Folder>에 출력하기
        # 폴더 이름 저장
        self.folder_name = self.folder_path.split('/')[-1]
        
        # show filelist  @ http://daplus.net/python-%ED%8C%8C%EC%9D%B4%EC%8D%AC%EC%9D%80-%EC%97%AC%EB%9F%AC-%ED%8C%8C%EC%9D%BC-%ED%98%95%EC%8B%9D%EC%9D%84-%EA%B0%80%EC%A0%B8%EC%98%B5%EB%8B%88%EB%8B%A4/ 
        FileList = [ x for x in os.listdir(self.folder_path) if re.match('.*\.JPG|.*\.jpg|.*\.PNG|.*\.png|.*\.TIF|.*\.TIFF|.*\.tif|.*\.tiff', x)] 
        FileList.sort()
        
        ## qtablewidget에 출력하기  # https://blog.naver.com/PostView.nhn?isHttpsRedirect=true&blogId=anakt&logNo=221834285100
        cell_row = len(FileList)
        col = 0
        self.filelist.setRowCount(cell_row)  # https://doc.qt.io/qt-5/qtablewidget.html
        
        for i in range(cell_row):
            FileName = str(FileList[i])  # 이미지파일 이름 (상대적 경로)
            
            item = QTableWidgetItem(FileName)
            self.filelist.setItem(i, col, item)

    def select_cell(self): 
        
        #### 초기화
        # self.box_ptrs.clear()
        self.box_dict.clear()
        # self.box_0_dict = {i:np.nan for i in range (1,9)}
        self.box_0_dict = {}
        self.box_vi_dict = {}
        # self.box_0_list.append(self.box_0_dict)
        self.box_id = 0  # 번호 초기화
        self.scene.clear()  # box 제거 -> 오류 : 모두 제거 ; https://doc.qt.io/archives/qt-4.8/qgraphicsscene.html#clear
        
        cell_Key = self.filelist.selectedIndexes()  # 리스트로 선택된 행번호와 열번호가 cell에 입력된다.

        ## 행/열 번호 콜
        self.cell_row = cell_Key[0].row()  # 선택된 행 번호 콜
        self.cell_col = cell_Key[0].column()  # 선택된 열 번호 콜

        ## 셀에 들어 있는 값을 가져오기
        self.file_name = self.filelist.item(self.cell_row, self.cell_col).text()  # 선택된 행과 열로 이루어진 텍스트값 변수화.
        self.file_path = self.folder_path + '/' + self.file_name  # 이미지파일 이름 (절대적 경로)   
           
        ## 엑셀 저장 포멧 초기화
        # self.box_0_dict['filename'] = self.file_name

        ## displayImg 함수 호출
        self.display_img()  # <Display_imgFile>함수 호출
        
    def display_img(self):
        
        ## 영어 경로만 가능
        # im = cv2.imread(self.file_path, cv2.IMREAD_UNCHANGED)     # cv2.imread로 바로 읽는 방법.
        ## 한글 경로 > NoneType Error : https://bskyvision.com/m/entry/python-cv2imread-%ED%95%9C%EA%B8%80-%ED%8C%8C%EC%9D%BC-%EA%B2%BD%EB%A1%9C-%EC%9D%B8%EC%8B%9D%EC%9D%84-%EB%AA%BB%ED%95%98%EB%8A%94-%EB%AC%B8%EC%A0%9C-%ED%95%B4%EA%B2%B0-%EB%B0%A9%EB%B2%95
        n = np.fromfile(self.file_path, np.uint8)   # binary data를 numpy 행렬로 읽음.
        im = cv2.imdecode(n, cv2.IMREAD_UNCHANGED)  # opencv에서 사용할 수 있는 형태로 바꿔 줌.
        
        if im.dtype !=np.uint8:  # vi
            ## convert float32 to uint8
            im[im<0]=0  # vi 음수값 0으로 채우기. 어차피 0-1 값만 의미 있기 때문.       
            im *= 255  #@times
            im = im.astype(np.uint8)
            
            ## set QPixmap : https://stackoverflow.com/questions/61910137/convert-python-numpy-array-to-pyqt-qpixmap-image-result-in-noise-image
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
        
        ''' [추가 기능 5] pixmap size fixing '''
        ## resize image to fit in graphics view : https://forum.qt.io/topic/92586/resize-image-to-fit-in-graphics-view
        self.view.fitInView(pixmapitem)
        
        ''' [추가 기능 5] pixmap size fixing '''   
               
    def create_box(self):
        
        ''' [추가 기능 1] box size update '''
        self.w_box = int(self.line_edit_box_w.text())
        self.h_box = int(self.line_edit_box_h.text())
        ''' [추가 기능 1] box size update '''
        
        ### box item 생성!
        self.box = QGraphicsRectItem(0, 0, self.w_box, self.h_box)  # 0으로 했는데 왜 -0.5 나오지..?
        # print('area : ', self.box.sceneBoundingRect().topLeft(), self.box.sceneBoundingRect().bottomRight())

        ''' [추가 기능 2] box rep update '''
        self.box_rep_num = int(self.line_edit_box_rep.text())
        self.draw_box_line(self.box_rep_num)
        ''' [추가 기능 2] box rep update '''        
               
        # define the brush
        brush = QBrush(Qt.cyan, Qt.Dense6Pattern)   # Qt.BrushStyle : https://doc.qt.io/archives/qtjambi-4.5.2_01/com/trolltech/qt/core/Qt.BrushStyle.html | https://wikidocs.net/37095
        self.box.setBrush(brush)
        
        # Draw a rectangle item, setting the dimensions.
        self.scene.addItem(self.box)
        
        # Set the self.box item as moveable and selectable.  @ https://stackoverflow.com/questions/27742557/pyqt-moving-multiple-items-with-different-itemignorestransformations-flags
        self.box.setFlag(QGraphicsItem.ItemIsMovable)
        self.box.setFlag(QGraphicsItem.ItemIsSelectable)        
        
        ## box_id
        ''' [추가 기능 3] box name update '''
        self.box_id += 1
        self.box_name = self.line_edit_box_id.text() + str(self.box_id)
        ''' [추가 기능 3] box name update '''     
        text = self.scene.addText('ID ' + str(self.box_name), QFont("Times", 11, QFont.Bold))  # Qfont : https://www.delftstack.com/ko/tutorial/pyqt5/pyqt5-label/
        text.setPos(self.box.pos().x(), self.box.pos().y()-20)
        text.setDefaultTextColor(QColor(Qt.cyan))   #html color chart : https://wepplication.github.io/tools/colorPicker/  # m2 : setHtml() : https://stackoverflow.com/questions/27612052/addtext-change-the-text-color-inside-a-qgraphicsview  
        text.setParentItem(self.box)  # Group box and box_name  @ https://www.qtcentre.org/threads/65316-QGraphicsText-parent-item-on-top-of-QGraphicsRectItem
                
        ## add box to box_0_dict
        degree = 0
        self.box_dict[self.box_name] = [degree, self.box]

    ''' [추가 기능 2] box rep update '''
    def draw_box_line(self, N:int):
        
        # print('N : ', N)
        
        # define the pen
        pen = QPen(Qt.cyan)
        pen.setWidth(2)
        self.box.setPen(pen)
        pen.setWidth(2)
        
        ## N등분 위해서 line item 생성
        for i in range(N):
            line = QGraphicsLineItem(self.w_box/N*(N-i),0, self.w_box/N*(N-i),self.h_box)
            line.setParentItem(self.box)
            line.setPen(pen)
        
        # Get the coordinates QGraphicsLineItem : https://stackoverflow.com/questions/37397864/get-the-coordinates-qgraphicslineitem
        # print('line1 : ', self.line1.line().p1().y(), self.line1.line().p2().y())  
        
    def delete_box(self):
        
        items = self.scene.selectedItems()
        for item in items:
            
            # 화면에서 없어지게
            self.scene.removeItem(item) 
            
            # box_id 전으로 초기화
            self.box_id -= 1  
    
    def RotateCoordinate(self, org_pt, theta, standart_pt):
        import math
    
        rot_x = (org_pt[0] - standart_pt[0])*math.cos(math.radians(theta)) - (org_pt[1] - standart_pt[1])*math.sin(math.radians(theta)) + standart_pt[0]

        rot_y = (org_pt[0] - standart_pt[0])*math.sin(math.radians(theta)) + (org_pt[1] - standart_pt[1])*math.cos(math.radians(theta)) + standart_pt[1]
        
        # print('standard pt : ', standart_pt)
        # print('org pt : ', org_pt)

        return rot_x, rot_y
    
    def rotate_box(self, value):
                
        items = self.scene.selectedItems()
        for item in items:
            
            ## box
            # item.transformOriginPoint()
            item.setRotation(value)

            box_dict = self.box_dict[self.box_id]
            self.box_dict[self.box_id] = [value, box_dict[1]]
            # print('box dict when rotated : ', self.box_dict[self.box_id][1].pos())
            # print('total box dict : ', self.box_dict)
            
            # bin = np.zeros((1700,1000), np.uint8)
            # mask = cv2.fillPoly(bin, [box1_ptrs], 255, cv2.LINE_AA)
            # cv2.imshow('mask', mask)
            # cv2.waitKey()
    
    '''[추가 기능 2] box rep update '''
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
    
    ## vi 범위만 가능!  # Result = average vi value                
    def save_vi_roi(self):
                
        ## box 넓이 -> 아! 고정이지! 영역 내 0 개수가 중요한 거 였지!!!
        # 다각형 넓이 구하는 공식(그린정리) - 사용하지 않았지만, 좋은 개념 : https://chocochip101.tistory.com/entry/%EB%B0%B1%EC%A4%80-2166%EB%B2%88-Python-%EB%8B%A4%EA%B0%81%ED%98%95%EC%9D%98-%EB%A9%B4%EC%A0%81

        ## 이미지에서 box 영역 0 구하기  @ https://copycoding.tistory.com/150
        im = cv2.imread(self.file_path, cv2.IMREAD_UNCHANGED)
        # im = cv2.resize(im, (self.view.width(), self.view.height()), cv2.INTER_LINEAR)   # https://076923.github.io/posts/Python-opencv-8/
        im = cv2.resize(im, (self.w, self.h), cv2.INTER_LINEAR)   # https://076923.github.io/posts/Python-opencv-8/
                   
        bin_total = np.zeros(im.shape, np.uint8) 
        
        ## Reset the dictionary
        self.box_vi_dict = {} 
        self.box_json_dict = {}  # [sol]
        self.box_json_dict['filename'] = self.file_name  # instead, from collections import OrderedDict
        
        for box_id, box_lines in self.box_dict.items():  # for key,value in dict.items() :t https://dojang.io/mod/page/view.php?id=2308
            
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
            # rot_ptrs_list.append([[rot_x5, rot_y5], [rot_x6, rot_y6]])  # 위 함수에서 반복되기 때문 (쓰고 싶으면, range (1,N))
            # print('rot_ptrs_list : ', rot_ptrs_list)
            
            box_ptrs_list = []
            for i in range(self.box_rep_num):
                
                box_ptrs_list.append(np.array([rot_ptrs_list[i][0], rot_ptrs_list[i+1][0], rot_ptrs_list[i+1][1], rot_ptrs_list[i][1]], np.int32))
                
            # print('box_ptrs_list : ', box_ptrs_list)     

            '''            
            ## 이미지에서 box 영역 0 구하기  @ https://copycoding.tistory.com/150       
            im = cv2.imread(self.file_path, cv2.IMREAD_UNCHANGED)  # vi
            im = cv2.resize(im, (self.view.width(), self.view.height()), cv2.INTER_LINEAR)   # https://076923.github.io/posts/Python-opencv-8/
            # im = cv2.resize(im, (self.w, self.h), cv2.INTER_LINEAR)   # https://076923.github.io/posts/Python-opencv-8/
            '''
            
            im[im<0] = 0   ### 최소값 0 : https://rfriend.tistory.com/426
            im[im>1] = 1   ### 최댓값 1 
                            
            vi_list = []
            for box_ptrs in box_ptrs_list:
                # print(box_ptrs)
                
                bin = np.zeros(im.shape, np.uint8)
                mask = cv2.fillPoly(bin, [box_ptrs], 255, cv2.LINE_AA)
                mask_total = cv2.fillPoly(bin_total, [box_ptrs], 255, cv2.LINE_AA)
                im_box = cv2.bitwise_and(im,im, mask=mask)
                im_box_total = cv2.bitwise_and(im,im, mask=mask_total)          
                
                # cv2.imshow('mask', mask)
                # cv2.imshow('im_box', im_box)
                # cv2.waitKey()
                
                # sum(vi)/sum(px) 구하기 
                vi_px = cv2.countNonZero(im_box)
                vi_sum1 = sum(im_box)   # https://codechacha.com/ko/python-sum/             
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
        
        # QMessageBox.information(self, "QMessageBox", "Successfully Saved the Data!")  # https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=huntingbear21&logNo=221825482802      

    ## show on table_db
    def setTableWidgetData(self, dict):
        
        # filename deletion
        # if 'filename' in dict:
        #     del(dict['filename'])
        #     print('dict : ', dict)
        
        # print('dict num :', len(dict))
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
                val = str(np.round(val, 5))  # 소수점 셋째자리 반올림
                # print('row, val : ', row, val)
                
                item_id = QTableWidgetItem(k)
                item_vi = QTableWidgetItem(val)

                self.table_db.setItem(row, 0, item_id)
                self.table_db.setItem(row, 1, item_vi)
            
    def save_img(self):
        
        # print('filename : ', self.file_name)
        try:
            # path
            save_img_path = self.save_dir + '/' + self.file_name
            
            # ScreenShot : https://kminseo.tistory.com/7
            p = QScreen.grabWindow(app.primaryScreen(), w.winId())#(메인화면, 현재위젯)
            p.save(save_img_path, 'jpg')
            
        except TypeError as e:
            # QMessageBox.critical(self, 'error', str(e), buttons=QMessageBox.Ok)
            QMessageBox.information(self, "Error", 'Make the Output Folder Path!')
                
    def cell_bg_color(self):
        
        ## M1
        # item = QTableWidgetItem(self.file_name)
        # self.filelist.setItem(self.cell_row, self.cell_col, item)
        # item.setBackground(Qt.cyan)
        
        ## M2
        self.filelist.item(self.cell_row, self.cell_col).setBackground(Qt.cyan)
    
    ## 이거 미소테크 답변 듯고 바꾸기. yaml or json,,,        
    def save_excel(self):
        
        # data=self.box_vi_list
        # print('data: ', self.data)
        
        df = pd.DataFrame(self.data)
        df.set_index('filename', inplace=True)   # filename 열을 인덱스로 지정
        
        # 저장 경로 선택
        save_exl_path = self.save_dir + '/' + self.folder_name + '.xlsx' 
        save_json_path = self.save_dir + '/' + self.folder_name + '.json' 
        
        ## Save VI as Excel file
        # 열 너비 자동 조절 : http://daplus.net/python-pandas-excelwriter%EB%A1%9C-excel-%EC%97%B4-%EB%84%88%EB%B9%84%EB%A5%BC-%EC%9E%90%EB%8F%99-%EC%A1%B0%EC%A0%95%ED%95%98%EB%8A%94-%EB%B0%A9%EB%B2%95%EC%9D%B4-%EC%9E%88%EC%8A%B5%EB%8B%88/
        writer = pd.ExcelWriter(save_exl_path, engine='xlsxwriter')
        df.to_excel(writer, na_rep = 'NaN', sheet_name=self.folder_name )
        writer.save()
        
        ## Save 
        df_json = self.json
        with open(save_json_path, "w") as json_file:
            json.dump(df_json, json_file)
            
        QMessageBox.information(self, "QMessageBox", "Successfully Saved All Datas as Excel!")  # https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=huntingbear21&logNo=221825482802
    
    def my_exception_hook(exctype, value, traceback):
        
        '''pyqt5 에서 exception 발생시 종료 방지 방법'''

        # show the error message box to notify to users
        if exctype == TypeError:
            # show the error message box
            error_message = "A TypeError occurred: " + str(value)
            # QMessageBox.critical(None, "Type Error", error_message, buttons=QMessageBox.Ok)

            # QMessageBox.information(None, "Error", error_message)
        else:   # call the default exception hook
            # Print the error and traceback
            # print(exctype, value, traceback)
            # Call the normal Exception hook after
            sys._excepthook(exctype, value, traceback)
            # sys.exit(1)     

    def closeEvent(self, QCloseEvent):
        
        '''종료이벤트 발생시 메세지박스로 종료 확인 함수'''  # [출처] [파이썬 라이브러리]#5 PyQt5 종료이벤트 추가|작성자 으뜸아빠
        
        # reQ = QMessageBox.question(self, "Check Exit", "Are you sure you want to quit the macro?",
        #             QMessageBox.Yes|QMessageBox.No)

        # if reQ == QMessageBox.Yes:
        #     QCloseEvent.accept()
        # else:
        #     QCloseEvent.ignore()


if __name__ == "__main__":

    app = QApplication(sys.argv)
    w = KUWindow()
    w.show()
    sys.exit(app.exec())
    
# error : 엑셀 파일 저장 안 됨. => pandas 설치된 것 확인 ; https://m.blog.naver.com/PostView.naver?isHttpsRedirect=true&blogId=ajdkfl6445&logNo=221351895331
# 파일 열리는 시간 너무 많음 =>
# qtablewidget : 스크롤바 되는지 => good