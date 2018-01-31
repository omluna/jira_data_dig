import time,os
from selenium import webdriver
from pyvirtualdisplay import Display

      
### from bokeh/io, slightly modified to avoid their import_required util
### didn't ultimately use, but leaving in case I figure out how to stick wtih phentomjs
### - https://github.com/bokeh/bokeh/blob/master/bokeh/io/export.py
def create_default_webdriver():
    '''Return phantomjs enabled webdriver'''
    return webdriver.PhantomJS(executable_path='~/phantomjs-2.1.1-linux-x86_64/bin/phantomjs', service_log_path=devnull)


### based on last SO answer above
### - https://stackoverflow.com/questions/38615811/how-to-download-a-file-with-python-selenium-and-phantomjs
def create_chromedriver_webdriver(dload_path):
    display = Display(visible=0)
    display.start()
    chrome_options = webdriver.ChromeOptions()
    prefs = {"download.default_directory": dload_path}
    chrome_options.add_experimental_option("prefs", prefs)
    driver = webdriver.Chrome(chrome_options=chrome_options)
    return driver, display


def download_file(dload):
    os.chdir(dload)
    html_file = [fn for fn in os.listdir() if fn.endswith('.html')]

    ### original code contained height/width for the display and chromium webdriver
    ### I found they didn't matter; specifying the image size to generate will 
    ### produce a plot of that size no matter the webdriver
        #py.plot(fig, filename=html_file, auto_open=False,
                        #image_width=1280, image_height=800,
                        #image_filename=fname, image='png')


    ### create webdrive, open file, maximize, and sleep
    driver, display = create_chromedriver_webdriver(dload)

    for filename in html_file:
        driver.get('file:///{}'.format(os.path.abspath(filename)))

        # make sure we give the file time to download
        time.sleep(2)

    ### was in the SO post and could be a more robust way to wait vs. just sleeping 1sec
    # while not(glob.glob(os.path.join(dl_location, filename))):
    #     time.sleep(1)

    driver.close()
    display.stop()
