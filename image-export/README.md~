__[PageNamesToStory_ConditionalText.jsx](pagenames/PageNamesToStory_ConditionalText.jsx)__

Will first identify the story that runs through the highest number of pages (called `mainStory` in the script).

Will then add conditional text to the beginning and end of each page in mainStory:

 * Condition: PageStart. Text: 'PageStart_[name]'.
 * Condition: PageEnd. Text: 'PageEnd_[name]'.
 
This information may then be processed by a converter that works on an IDML export, since it will otherwise not know on which page which content is placed.

Will remove existing PageStart/PageEnd conditional text before inserting new PageStart/PageEnd conditional text.

In order to get accurate information, make sure to process the document with all relevant fonts installed so that the pagination is the same.

If there are any objects between a PageEnd and the subsequent PageStart label, a processor will not know on which page they are exactly. It might therefore be important to anchor tables or figures with captions before the last mainStory character on a page, because otherwise the PageEnd label might be inserted before the anchor.

This script will not work well for documents that don’t have a main story, i.e., where text does not flow from frame to frame across the pages or where each chapter is a story of its own. There will be another version or an option to better support these use cases.
