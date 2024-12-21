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

        # create presentation object
        prs = pp.Presentation()

        # TEST
        # add a slide
        slide_layout = prs.slide_layouts[self.TITLE]
        slide1 = prs.slides.add_slide(slide_layout)
        # put stuff on the slide
        shapes = slide1.shapes

        # create title slide

        # while csv interpreter still has slides to add loop through a slide adding method

        # save the presentation
        prs.save("pptx_exports/" + cli_instance.get_PowerPoint_SaveName() + ".pptx")