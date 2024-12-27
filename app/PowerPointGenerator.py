import pptx as pp
from pptx.util import Inches

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

    def __init__(self, cli_instance):
        print("Starting PowerPoint Generation...")
        self._cli = cli_instance

        # create presentation object
        self.prs = pp.Presentation()

        # create title slide
        slide_layout = self.prs.slide_layouts[self.TITLE]
        slide = self.prs.slides.add_slide(slide_layout)

        # put stuff on the slide
        placeholders = slide.placeholders

        title = placeholders[0]
        subtitle = placeholders[1]

        title.text = self._cli.get_PowerPoint_Name()
        subtitle.text = "Catholic Charities of East Tennessee"

    def create_Table_Slide(self, title, matrix, columns, rows):
        print ("\n" + "Columns: " + str(columns) + " | " + "Rows: " + str(rows) + "\n" + "Slide title: " + title)
        for row in matrix:
            print(row)

        slide_layout = self.prs.slide_layouts[self.TITLE_ONLY]
        slide = self.prs.slides.add_slide(slide_layout)
        shapes = slide.shapes
        shapes.title.text = title

        left = top = Inches(2.0)
        width = Inches(6.0)
        height = Inches(0.8)

        table = shapes.add_table(rows, columns, left, top, width, height).table

        # Set column widths
        #table.columns[0].width = Inches(2.0)

        # fill in cells
        for i in range(len(matrix)): # loop through columns
            for j in range(len(matrix[i])): # loop through rows
                table.cell(i, j).text = matrix[i][j]

    @staticmethod
    def create_PieChart_Slide(title, matrix):
        print("TBD")

    @staticmethod
    def create_BarChart_slide(title, matrix):
        print("TBD")

    def save_Presentation(self):
        # save the presentation
        self.prs.save("pptx_exports/" + self._cli.get_PowerPoint_SaveName() + ".pptx")