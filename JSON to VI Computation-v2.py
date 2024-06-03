# -*- coding: utf-8 -*-
import argparse
import cv2
import os, sys
import numpy as np
import pandas as pd
import json
   
   
W, H = 1150, 880  # resize
                
def save_excel(data, save_dir, folder_name):
    
    # print('data: ', data)
    
    # Convert the list to a DataFrame  # dict: df = pd.DataFrame.from_dict(my_dict, orient='index')
    df = pd.DataFrame(data)
    df.set_index('filename', inplace=True)# filename 열을 인덱스로 지정
    
    # Excel file path
    save_exl_path = save_dir + '/' + folder_name + '.xlsx' 
    
    ## Save VI as Excel file
    # Automated adjustment in column width
    writer = pd.ExcelWriter(save_exl_path, engine='xlsxwriter')
    # Export the DataFrame to Excel without square brackets
    df.to_excel(writer, na_rep = 'NaN', sheet_name=folder_name )
    writer.save()

date = '20240603'
json_path = os.path.join('sample/nir_db',date,date+'.json')
in_dir = os.path.join('sample/ndvi', date)
filelist = os.listdir(in_dir)
filelist.sort()
out_dir = os.path.join('sample/ndvi_db', date)
os.makedirs(out_dir, exist_ok=True)

# import JSON file 
with open(json_path, 'r') as f:
    data = json.load(f)  # json file
    # print('data:', data)

# Computate the RoI in the VI images
num = 0
box_ndvi_list = []
for dic in data:
    
    print('num : ', num)
    
    # filename
    filename = dic['filename']
    vi_filename = '_'.join(filename.split('_')[:2])+'_NDVI.TIF'
    print('filename : ', vi_filename)
    num += 1
    
    del(dic['filename'])
    box_ptrs_dic = dic
    
    # Read the file
    file_path = os.path.join(in_dir, vi_filename)
    print('file_path:', file_path)
    im = cv2.imread(file_path, cv2.IMREAD_UNCHANGED)
    im = cv2.resize(im, (W, H), cv2.INTER_LINEAR)   

    # Reset the dictionary 
    box_ndvi_dict = {} 

    im[im<0] = 0   ### 최소값 0 : 
    im[im>1] = 1   ### 최댓값 1                  
    bin_total = np.zeros(im.shape, np.uint8)      
    # for box_ptrs in box_ptrs_list:
    for box_id, box_ptrs_list in box_ptrs_dic.items():
        ndvi_list = []        
        for box_ptrs in box_ptrs_list:
            
            # list to array
            box_ptrs = np.array([box_ptrs])
            
            bin = np.zeros(im.shape, np.uint8)
            mask = cv2.fillPoly(bin, box_ptrs, 255)
            mask_total = cv2.fillPoly(bin_total, [box_ptrs], 255, cv2.LINE_AA)
            im_box = cv2.bitwise_and(im,im, mask=mask)
            im_box_total = cv2.bitwise_and(im,im, mask=mask_total)                
            
            # sum(ndvi)/sum(px) 구하기 
            ndvi_px = cv2.countNonZero(im_box)
            ndvi_sum1 = sum(im_box)              
            ndvi_sum = sum(ndvi_sum1)
            ndvi = ndvi_sum/ndvi_px
            ndvi_list.append(ndvi)      
        box_ndvi_dict[box_id]= ndvi_list
    
    out_path = os.path.join(out_dir, vi_filename)
    cv2.imwrite(out_path, im_box_total)
        
    # remove the square brackets []
    box_ndvi_dict = {key: str(value).strip('[]') for key, value in box_ndvi_dict.items()}
            
    box_ndvi_dict['filename'] = vi_filename

    box_ndvi_list.append(box_ndvi_dict) 
    # break
               
# Save the data
data = box_ndvi_list
save_excel(data, out_dir, date)
    

        

# %%
