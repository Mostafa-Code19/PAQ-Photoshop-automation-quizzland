import os, time
from photoshop import Session

def App(inputDir, targetImg, userChoseForFinalWidth):
    with Session() as ps:
        # open the img and ps
        ps.app.documents.add(1920, 1080, name="content")
        desc = ps.ActionDescriptor
        desc.putPath(ps.app.charIDToTypeID("null"), f"{inputDir}/{targetImg}")
        event_id = ps.app.charIDToTypeID("Plc ")  # `Plc` need one space in here.
        ps.app.executeAction(ps.app.charIDToTypeID("Plc "), desc)


        # trim
        ps.active_document.trim(ps.TrimType.TopLeftPixel, True, True, True, True)

        # input width of final product
        if userChoseForFinalWidth == 1:
            imgFinalWidth = 1336
            imgFinalHeight = 768

        elif userChoseForFinalWidth == 2:
            imgFinalWidth = 625
            imgFinalHeight = 359

        elif userChoseForFinalWidth == 3:
            imgFinalWidth = 520
            imgFinalHeight = 624
        else:
            print('You Input Is Invalid. Please Try Again!')
            # return App(inputDir, targetImg, userChoseForFinalWidth)
            
        # resize
        ps.active_document.resizeImage(imgFinalWidth, None)

        widthOfImg = ps.active_document.Width
        heightOfImg = ps.active_document.Height

        # select
        def selectAndCrop():
            resultWidthByRatio9 = (heightOfImg * 16) / 9
            isRatio16By9 = widthOfImg - 100 <= resultWidthByRatio9  <= widthOfImg + 100

            start_ruler_units = ps.app.preferences.rulerUnits
            if start_ruler_units is not ps.Units.Pixels:
                ps.app.Preferences.RulerUnits = ps.Units.Pixels

            if userChoseForFinalWidth != 3:
                if isRatio16By9:
                    # print('16:9')
                    return (
                        ps.active_document.crop(bounds=[0, 0, widthOfImg, heightOfImg], width=imgFinalWidth, height=imgFinalHeight)
                    )

                else:
                    # print('un16:9')
                    heightWithRatio16By9 = (widthOfImg * 9) / 16
                    offsetHeight = heightWithRatio16By9 / 9
                    # widthWithRatio16By9 = (heightOfImg * 16) / 9
                    # offsetWidth = widthWithRatio16By9 / 16

                    return (
                        ps.active_document.crop(bounds=[0, offsetHeight, widthOfImg, heightWithRatio16By9 + offsetHeight], width=imgFinalWidth, height=imgFinalHeight)
                    )
                    
            else:
                WidthWithRatio7By9 = (heightOfImg * 7.5) / 9
                offset = (widthOfImg / 7.5) * 2

                return (
                    ps.active_document.crop(bounds=[offset, 0, WidthWithRatio7By9 + offset, heightOfImg], width=imgFinalWidth, height=imgFinalHeight)
                )

        # crop
        ps.active_document.selection.select(selectAndCrop())

        # save
        saveDirection = inputDir + '/output'
        dir = os.path.join(saveDirection, str(time.time()))
        ps.active_document.saveAs(dir, ps.JPEGSaveOptions())


inputDir = str(input('\nEnter the directory: \n')).replace("\\", '/')
userChoseForFinalWidth = int(input('\n1:Background 2:Thumbnail+Question 3:QuizOption : '))

imgFiles = os.listdir(inputDir)
totalCountImgFiles = len(imgFiles) - 1  # minus 1 for output file
imgCounter = 1

print('Setup the program...')
try:
    for targetImg in imgFiles:
        if '.' in targetImg:  # if file skip
            App(inputDir, targetImg, userChoseForFinalWidth)

        print(f'{imgCounter} / {totalCountImgFiles} Done ✅', end="\r")
        imgCounter += 1
    print('-------------------------- Complete ✅ ----------------------------')

except Exception as e:
    print('There is not image in this direction \nPlease try again!')

input('Press Enter To Exit...')