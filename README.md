# Solution to WLODEK JSON to word task

## What it does
    It generates the titles and table of contents using appropriate numbering:
    HEADING_1 for Section Introduction
    HEADING_2 for Title
    HEADING_3 for Key Leadning Points

# How to use
   1. Clone the repo
   2. Run npm install
   3. Run node index.js
   4. My Document.docx gets created in the working directory


# telebrief.json build
 - File telebrief.json contains course build of the Telebrief course.
 - There are slight changes to course build .json
 1. DisplayTitle is now displayTitle
 2. Components contain properties object
 3. type prop is prefixed with underscore _type
 