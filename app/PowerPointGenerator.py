"""
File:       PowerPointGenerator.py
Purpose:    This file PowerPointGenerator.py contains the PowerPointGenerator class, which is responsible for the
            creation of the PowerPoint (slides, charts, graphs).
Author:     Joey Borrelli, Software & Training Intern For Catholic Charities of East Tennessee
Anno:       Anno Domini 2024
"""

import pptx as pp
from pptx.util import Inches
from pptx.enum.text import PP_ALIGN
import CLI as UI

class PowerPointGenerator:
    # Slide layout constants
    TITLE = 0
    TITLE_AND_CONTENT = 1
    SECTION_HEADER = 2
    TWO_CONTENT = 3 # side by side bullet textboxes
    COMPARISON = 4 # side by side bullet textboxes with titles
    TITLE_ONLY = 5
    BLANK = 6
    CONTENT_CAPTION = 7
    PICTURE_CAPTION = 8

    def __init__(self):
        print("Starting PowerPoint Generation...")

        # create presentation object
        self.prs = pp.Presentation()

        # create title slide
        slide_layout = self.prs.slide_layouts[self.TITLE]
        slide = self.prs.slides.add_slide(slide_layout)

        # put stuff on the slide
        placeholders = slide.placeholders

        title = placeholders[0]
        subtitle = placeholders[1]

        title.text = UI.get_PowerPoint_Name()
        subtitle.text = "Catholic Charities of East Tennessee"

    def create_Table_Slide(self, title, matrix, columns, rows):
        #print ("\nSlide title: " + title + "\n" + "Columns: " + str(columns) + " | " + "Rows: " + str(rows))
        #for row in matrix:
            #print(row)
        if rows > 0 and columns > 0:
            slide_layout = self.prs.slide_layouts[self.TITLE_ONLY]
            slide = self.prs.slides.add_slide(slide_layout)
            shapes = slide.shapes
            shapes.title.text = title

            left = Inches(0.0)
            top = Inches(2.0)
            width = Inches(10.0)
            height = Inches(0.8)

            table = shapes.add_table(rows, columns, left, top, width, height).table

            # Set column widths
            #table.columns[0].width = Inches(2.0)

            # fill in cells with data
            for row in range(len(matrix)): # loop through rows
                for col in range(len(matrix[row])): # loop through columns
                    if matrix[row][col] == 'Count':
                        matrix[row][col] = ''
                    table.cell(row, col).text = matrix[row][col]

            # merge the first row's cells
            table.cell(0, 0).merge(table.cell(0, columns - 1))
            # center text in merged cell
            for paragraph in table.cell(0, 0).text_frame.paragraphs:
                paragraph.alignment = PP_ALIGN.CENTER  # Horizontal alignment
            table.cell(0, 0).vertical_alignment = "middle"  # Vertical alignment

            # merge any row's cells where the first element isn't '', but every element after is
            for row in range(len(matrix)):
                if matrix[row][0] != '' and all(cell == '' for cell in matrix[row][1:]):
                    table.cell(row, 0).merge(table.cell(row, columns - 1))
        else:
            print("\nError creating slide " + title + ", rows or columns are < 0")

    @staticmethod
    def create_PieChart_Slide(title, matrix):
        print("TBD")

    @staticmethod
    def create_BarChart_slide(title, matrix):
        print("TBD")

    def save_Presentation(self):
        # save the presentation
        self.prs.save("pptx_exports/" + UI.get_PowerPoint_SaveName() + ".pptx")