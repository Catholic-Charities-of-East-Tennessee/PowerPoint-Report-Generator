import pptx as pp

class PowerPointGenerator:
    def __init__(self):
        print("Starting PowerPoint Generation...")
        # create presentation object
        pres = pp.Presentation()
        # create a layout object
        title_slide_layout = pres.slides[0]
        # create a slide object & add it to the presentation object
        title_slide = pres.slides.add_slide(title_slide_layout)
        # create a shape object off of the slide object
        title = title_slide.shapes.title
        # create another shape object off of the default placeholder list in the slide object
        subtitle = title_slide.placeholders[1]

        # set the text of both of the shape objects
        title.text = "Title"
        subtitle.text = "Subtitle"

        # save the presentation
        pres.save('pptx_exports/test.pptx')