import pptx as pp

class PowerPointGenerator:
    def __init__(self):
        print("Starting PowerPoint Generation...")
        my_presentation = pp.Presentation()
        my_presentation.save('test.pptx')