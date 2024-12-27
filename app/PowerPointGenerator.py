import pptx as pp

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
        slide1 = self.prs.slides.add_slide(slide_layout)

        # put stuff on the slide
        placeholders = slide1.placeholders

        title = placeholders[0]
        subtitle = placeholders[1]

        title.text = self._cli.get_PowerPoint_Name()
        subtitle.text = "Catholic Charities of East Tennessee"

    def create_Table_Slide(self, title, matrix):
        print ("\n" + title)
        for row in matrix:
            print(row)

        self.save_Presentation()

    def create_PieChart_Slide(self, title, matrix):
        print("TBD")
        self.save_Presentation()

    def create_BarChart_slide(self, title, matrix):
        print("TBD")
        self.save_Presentation()

    def save_Presentation(self):
        # save the presentation
        self.prs.save("pptx_exports/" + self._cli.get_PowerPoint_SaveName() + ".pptx")