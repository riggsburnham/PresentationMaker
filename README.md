# PresentationMaker
WPF Powerpoint Presentation Maker

A limitation for this project is that you can only add up to four pictures.

Can bold, italic and underline both title and description.

To bold use the hot key ctrl + b

To italic use the hot key ctrl + i

To underline use the hot key ctrl + u

In order to use application:
  1) Goto "https://developers.google.com/custom-search/v1/overview"
  2) Click link under Search engine ID labeled as "Programmable Search Engine control panel"
  3) Create search engine
  4) Select newly created search engine
  5) Make sure "Image Search" is turned on
  6) Make sure "Search the entire web" is turned on
  7) Copy and paste your search engine ID into the string "GOOGLE_IMAGE_SEARCH_API_URI" replacing the words "SEARCHENGINEIDGOESHERE" located in the GoogleImages class in the            GoogleLibrary project.
  8) Go back to "https://developers.google.com/custom-search/v1/overview"
  9) Click button under API Key labeled as "Get Key"
  10) Use drop down to select your newly created project, click next
  11) Copy and paste your API key into the string "GOOGLE_IMAGE_SEARCH_API_URI" replacing the words "KEYGOESHERE" located in the GoogleImages class in the GoogleLibrary project.
  12) If the steps above were done correctly the project is now ready to use
