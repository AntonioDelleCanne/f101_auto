import pandas as pd
import os
import sys
from docx import Document
from pathlib import Path
from nbdev.showdoc import doc
import ipywidgets as widgets
from ipywidgets import FileUpload
import numpy as np
import xlsxwriter
from openpyxl import Workbook
from openpyxl import load_workbook
from IPython.display import display
from ipywidgets import HTML
from functools import partial
import openpyxl
import os.path
import base64
import abc

from database import *

class UIObject:

    def __init__(self, output, reset_ui, db):
        self.output = output
        display(output)
        self.ui = []
        self.reset_ui = reset_ui
        self.db = db
        self.db.listen(self)
        
    
    def display(self):
        self.reset_ui()
        for d_data in self.ui:
            self.output.append_display_data(d_data)
    
    def notify(self):
        self.reset_ui()

class UploadButton(UIObject):
    '''
    Button to upload a file to root/fname
    '''
    def __init__(self, db, f_path=template_path, accept='.xlsx',output=widgets.Output()):
        super().__init__(output, self.reset_ui, db)
        
        self.f_path = f_path
        if not os.path.isfile(self.f_path):
            raise ValueError('File does not exist')
        
        self.uploader = FileUpload(accept='.xlsx', multiple=False)
        self.submit = widgets.Button(description='Submit')
        self.submit.on_click(self.__on_submit)
        self.msg = widgets.HTML(
            value=""
        )
        self.ui = [self.uploader, self.submit, self.msg]
    
    def reset_ui(self):
        self.uploader._counter = 0
        self.msg.value = ''
        if len(self.uploader.value.keys()) > 0:
            del self.uploader.value[list(self.uploader.value.keys())[0]]
        

    def __on_submit(self, b):
        if(self.uploader._counter == 0):
            self.msg.value= 'Please upload a file.'
            return
        with open(self.f_path, 'wb') as f:
            file = self.uploader.value
            file_b = list(file.values())[0]['content']
            f.write(file_b)
        self.reset_ui()
        self.msg.value = 'File submitted successfully!'

class DownloadButton(UIObject):
    '''
    Button to download a file to root/fname
    '''
    def __init__(self, db, to_download_path = download_zip, output=widgets.Output()):
        super().__init__(output, self.reset_ui, db)
        self.f_name = to_download_path.name
        self.f_path = to_download_path
        if not os.path.isfile(self.f_path):
            raise ValueError('File does not exist')
        self.button = widgets.HTML(
            value=""
        )
        self.ui = [self.button]
        
    def reset_ui(self):
        with open(self.f_path, 'rb') as f:
            content = f.read()
            b64 = base64.b64encode(content)
            payload = b64.decode()

        self.button.value = '''<html>
        <head>
        <meta name="viewport" content="width=device-width, initial-scale=1">
        </head>
        <body>
        <a download="{filename}" href="data:text/csv;base64,{payload}" download>
            <button class="p-Widget jupyter-widgets jupyter-button widget-button">Download Excel</button>
        </a>
        </body>
        </html>
        '''.format(payload=payload,filename=self.f_name)
            
class SubmitDocxUI(UIObject):
    '''
    UI to submit the feedback form
    '''
    def __init__(self, db, output=widgets.Output()):
        super().__init__(output, self.reset_ui, db)
        self.fname_text = widgets.Text(
            value='',
            placeholder='First Name',
            description='First Name:',
            disabled=False
        )
        self.lname_text = widgets.Text(
            value='',
            placeholder='Last Name',
            description='Last Name:',
            disabled=False
        )
        
        self.dropdown = widgets.Dropdown(
            options=[''],
            value='',
            description='Course Title:',
            disabled=False,
        )
        self.uploader = FileUpload(accept='.docx', multiple=False)
        self.button = widgets.Button(description='Submit')
        self.msg = widgets.HTML(
            value=""
        )
        self.button.on_click(self.__on_submit)
        self.ui = [self.fname_text, self.lname_text, self.dropdown, self.uploader, self.button, self.msg]
    
    def reset_ui(self):
        self.uploader._counter = 0
        self.msg.value = ''
        if len(self.uploader.value.keys()) > 0:
            del self.uploader.value[list(self.uploader.value.keys())[0]]
        course_titles = ['']
        course_titles = course_titles + self.db.get_courses()
        course_titles.append('Other')
        self.dropdown.value = ''
        self.dropdown.options = course_titles
        self.fname_text.value = ''
        self.lname_text.value = ''

        
    
    def __on_submit(self, b):
        name = self.fname_text.value + ' ' + self.lname_text.value
        name = name.lower()
        course_title = self.dropdown.value
        if not (bool(self.fname_text.value) and bool(self.lname_text.value) and bool(course_title) and bool(self.uploader.value)):
            self.msg.value = 'Please fill all the fields.'
        else:
            self.db.submit_docx(name, course_title, self.uploader, year=2020) # TODO change year
            self.reset_ui()
            self.msg.value = "File submitted successfully!"
                

class AddCourseUI(UIObject):
    def __init__(self, db, output=widgets.Output()):
        super().__init__(output, self.reset_ui, db)
        self.course_name_text = widgets.Text(
            value='',
            placeholder='New Course Name',
            description='New Course',
            disabled=False
        )
        self.button = widgets.Button(description='Submit')
        self.msg = widgets.HTML(
            value=""
        )
        self.button.on_click(self.__on_submit)
        self.ui = [self.course_name_text, self.button, self.msg]
        
    def reset_ui(self):
        self.course_name_text.value = ''
        self.msg.value = ''

        
    
    def __on_submit(self, b):
        v = str(self.course_name_text.value)
        if not bool(v):
            self.msg.value = 'Please fill all the fields.'
        elif v in self.db.get_courses():
            self.msg.value = 'Course already exists.'
        else:
            v = str(self.course_name_text.value)
            self.db.add_course(v)
            self.reset_ui()
            self.msg.value = f'Course "{v}" has been added.'
                
    
class DeleteCourseUI(UIObject):
    def __init__(self, db, output=widgets.Output()):
        super().__init__(output, self.reset_ui, db)
        self.dropdown = widgets.Dropdown(
            options=[''],
            value='',
            description='Course Title:',
            disabled=False,
        )
        self.button = widgets.Button(description='Delete')
        self.msg = widgets.HTML(
            value=""
        )
        self.button.on_click(self.__on_submit)   
        self.ui = [self.dropdown, self.button, self.msg]
        
    def reset_ui(self):
        course_titles = ['']
        course_titles = course_titles + self.db.get_courses()
        self.dropdown.value = ''
        self.dropdown.options = course_titles
        
    
    def __on_submit(self, b):
        v = str(self.dropdown.value)
        if not bool(v):
            self.msg.value= 'Please select a course.'
        else:
            self.db.remove_course(v)
            self.reset_ui()
            self.msg.value= f'Course "{v}" has been deleted.'
            
class ReviewFormsUI(UIObject):
    def __init__(self, db, output=widgets.Output()):
        super().__init__(output, self.reset_ui, db)
        self.dropdown_course = widgets.Dropdown(
            options=[''],
            value='',
            description='Course Title:',
            disabled=False,
        )
        
        self.dropdown_review = widgets.Dropdown(
            options=[''],
            value='',
            description='To Review:',
            disabled=False
        )
        
        self.submit = widgets.Button(description='Submit')
        self.submit.on_click(self.__on_submit)   
        
        self.msg = widgets.HTML(
            value=""
        )
        
        self.ui = [self.dropdown_course, self.dropdown_review, self.submit, self.msg]
        
    def reset_ui(self):
        course_titles = ['']
        course_titles = course_titles + self.db.get_courses()
        to_review = ['']
        to_review = to_review + self.db.get_review()
        self.dropdown_review.value = ''
        self.dropdown_review.options = to_review
        self.dropdown_course.value = ''
        self.dropdown_course.options = course_titles
        self.msg.value = ''
    
    def __on_submit(self, b):
        course = str(self.dropdown_course.value)
        review = str(self.dropdown_review.value)
        if not (bool(course) and bool(review)):
            self.msg.value= 'Please select all the fields.'
        else:
            self.db.review(review, course)
            self.reset_ui()
            self.msg.value= f'Submission "{review}" has been assigned to course "{course}".'
        