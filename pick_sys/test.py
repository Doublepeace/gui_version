#!/usr/bin/env python3

import pygal

def showBar():
    # ---------------------------------------------
    #chart = pygal.Line()
    fruitNo = [45, 52, 60, 72, 62, 35, 11,22,33,44,55]
    testno = [5, 2, 6, 7, 2, 5, 1,2,3,4,5]
    hist = pygal.Bar()
    hist.title = "熱門水果排行榜"
    hist.x_labels = ["芭樂", "柳丁", "蘋果", "鳳梨", "橘子", "西瓜","A", "b", "c","d","e"]
    hist.x_title = "水果種類"
    hist.y_title = "百人"
    hist.add("數量", fruitNo)
    hist.add("test", testno)
    hist.render_in_browser()#render_to_file("test.svg")


    #hist.render_to_png("test.png")


showBar()





