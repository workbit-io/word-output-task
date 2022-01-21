1. Set 3 global variables 
    nextHeading1num = 1;
    nextHeading2num = 1;
    nextHeading3num = 1;

2. Set variable for previous article (to check if previous was not title as if it was next will be a teaching point)
    previousArticle

Iterate over contentObject(each child is an object type: article)
if title = Section Introduction {
    go to children {
        iterate over children {
                if label = PowerPoint Image {
                get DisplayTitle": "Image Element Title",
                set nextHeading1num = 2;
                set previousArticle = sectionIntroduction;
                generate HEADING 1
            }
        }
    }
}
if title = title {
    go to children {
        go to children {
            get "DisplayTitle": "List Element TitleW",
            set nextHeading2num = 2;
            set previousArticle = title;
            generate HEADING 2
        }
    }
}
if title != Section introduction && title {         // it means it's a key learning point article
    iterate over Children(blocks - keyLearningPoints) {
        go to Children[0] {
            get "DisplayTitle": "How to take over the world",
            set nextHeading3num = 2;
            set previousArticle = teachingPoint;
            generate HEADING 3
        }
    }
}