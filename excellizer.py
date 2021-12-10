import xlsxwriter
from PIL import Image
from xlsxwriter.workbook import Workbook

class ExcelLizer:
    """
    A class that is able to create a preformatted Excel Sheet that can serve as a pixel canvas, and can take in a photo to place on that canvas.
    """
    def __init__(self):
        """
        Initializes the class with a default name 'Image' for the image that gets processed.
        """
        self.imageFileName = 'Image'

    def __init__(self, imageFileName: str):
        """
        Initializes the class with a parameter for the location where the photo to be processed is located as a String.
        """
        self.imageFileName = imageFileName
    
    def setImageFileName(self, name: str):
        """
        Allows for the location of the image file to be set after the initialization of the class.
        """
        self.imageFileName = name

    def createExcelCanvas(self, location: str='ExcelLizerBase.xlsx', imageName: str='Image', height: int=2826, width: int=4248) -> Workbook:
        """
        Creates an excel sheet with a pre-set conditional formatting style applied such that each group of 3 cells that are vertically aligned
        will have an RGB-color style similar to a digital pixel.
        If no dimensions are passed in for the sheet, it will default to generating a sheet which can handle up to a 12 megapixel photo.

        Returns the workbook that canvas was generated in.
        """
        workbook = xlsxwriter.Workbook(location)

        lastName = imageName.split("/")[-1]
        canvas = workbook.add_worksheet(lastName)
        # canvas.hide_gridlines(2)

        redFormat = {'type': '2_color_scale', 
        'min_type': 'num',
        'max_type': 'num',
        'min_value': 0,
        'max_value': 255,
        'min_color': '#000000'}

        greenFormat = redFormat.copy()
        blueFormat = redFormat.copy()

        redFormat['max_color'] = '#FF0000'
        greenFormat['max_color'] = '#00FF00'
        blueFormat['max_color'] = '#0000FF'

        for row in range(height * 3):
            canvas.set_row_pixels(row, 8)

            if row % 3 == 0: #red pixel
                canvas.conditional_format(row, 0, row, width, redFormat)
            elif row % 3 == 1: #green pixel
                canvas.conditional_format(row, 0, row, width, greenFormat)
            else: #blue pixel
                canvas.conditional_format(row, 0, row, width, blueFormat)

        for column in range(width):
            canvas.set_column_pixels(column, column, 24)
        
        return workbook
    
    def convertImage(self):
        """
        Converts the image into an Excel Sheet format. Creates the canvas excel sheet if it does not exist already.
        """
        print('Fetching Image...')
        image = Image.open(self.imageFileName).convert('RGB')
        width, height = image.size

        if width > 240 or height > 240:
            print('Downscaling Image...')
            image.thumbnail((240, 240), Image.ANTIALIAS)
            width, height = image.size

        print('Creating Workbook...', end=' ')

        lastName = self.imageFileName.split("/")[-1]
        workbook = self.createExcelCanvas(f'{lastName}-excelized.xlsx', self.imageFileName, height, width)
        
        print('Workbook formatted.')
        canvas = workbook.get_worksheet_by_name(lastName)

        print('Generating Image...', end=' ')
        for x_pixel in range(width):
            for y_pixel in range(height):
                r, g, b = image.getpixel((x_pixel, y_pixel))

                rowNum = y_pixel * 3
                canvas.write(rowNum, x_pixel, r, )
                canvas.write(rowNum + 1, x_pixel, g)
                canvas.write(rowNum + 2, x_pixel, b)
            
        print('Image Generated.')
        workbook.close()


if __name__ == '__main__':
    fileName = input('Enter file name of image: ')
    excellizer = ExcelLizer(fileName)
    excellizer.convertImage()


