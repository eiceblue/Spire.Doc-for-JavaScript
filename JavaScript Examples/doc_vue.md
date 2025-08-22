# spire.doc javascript hello world
## create a simple Word document with Hello World text
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Create a new paragraph
let paragraph = section.AddParagraph();

// Append text in paragraph
paragraph.AppendText('Hello World!');

// Save the document to the specified path
document.SaveToFile({
  fileName: 'HelloWorld.docx',
  fileFormat: wasmModule.FileFormat.Docx2013,
});

// Clean up resources
document.Dispose();
```

---

# Word Document Text Finding and Highlighting
## This code demonstrates how to find specific text in a Word document and highlight all occurrences
```javascript
// Find all occurrences of the string "word" in the document
let textSelections = doc.FindAllString('word', false, true);

// Iterate through all found text selections
for (let i = 0; i < textSelections.length; i++) {
  let selection = textSelections[i];

  // Set the highlight color of the selected text to yellow
  selection.GetAsOneRange().CharacterFormat.HighlightColor = wasmModule.Color.get_Yellow();
}
```

---

# Document Content Replacement
## Replace content in a document with another document using regex pattern matching
```javascript
//Create the first document
let document1 = wasmModule.Document.Create();

//Load the first document from disk.
document1.LoadFromFile(inputFileName1);

//Create the second document
let document2 = wasmModule.Document.Create();

//Load the second document from disk.
document2.LoadFromFile(inputFileName2);

//Get the first section of the first document
let section1 = document1.Sections.get(0);

//Create a regex
let regex = wasmModule.Regex.Create('\\[MY_DOCUMENT\\]', wasmModule.RegexOptions.None);

//Find the text by regex
let textSections = document1.FindAllPattern({ pattern: regex });

//Travel the found strings
for (let i = 0; i < textSections.length; i++) {
  let seletion = textSections[i];
  //Get the para
  let para = seletion.GetAsOneRange().OwnerParagraph;
  //Get textRange
  let textRange = seletion.GetAsOneRange();
  //Get the para index
  let index = section1.Body.ChildObjects.IndexOf(para);
  //Insert the paragraphs of document2
  for (let i = 0; i < document2.Sections.Count; i++) {
    let section2 = document2.Sections.get_Item(i);
    for (let j = 0; j < section2.Paragraphs.Count; j++) {
      let paragraph = section2.Paragraphs.get_Item(j);
      section1.Body.ChildObjects.Insert(index, paragraph.Clone());
    }
  }
  //Remove the found textRange
  para.ChildObjects.Remove(textRange);
}
```

---

# spire.doc javascript regex replacement
## replace text using regular expressions in a word document
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

//create a regex, match the text that starts with #
let regex = wasmModule.Regex.Create('\\#\\w+\\b', wasmModule.RegexOptions.None);

//replace the text by regex
doc.Replace(regex, 'Spire.Doc');

// Clean up resources
doc.Dispose();
```

---

# Spire.Doc JavaScript Text Replacement
## Replace text with field in a Word document
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

//Find the target text
let selection = doc.FindString({
  stringValue: 'summary',
  caseSensitive: false,
  wholeWord: true,
});
//Get text range
let textRange = selection.GetAsOneRange();
//Get it's owner paragraph
let ownParagraph = textRange.OwnerParagraph;
//Get the index of this text range
let rangeIndex = ownParagraph.ChildObjects.IndexOf(textRange);
//Remove the text range
ownParagraph.ChildObjects.RemoveAt(rangeIndex);
//Remove the objects which are behind the text range
let tempList = [];
for (let i = rangeIndex; i < ownParagraph.ChildObjects.Count; i++) {
  //Add a copy of these objects into a temp list
  tempList.push(ownParagraph.ChildObjects.get(rangeIndex).Clone());
  ownParagraph.ChildObjects.RemoveAt(rangeIndex);
}
//Append field to the paragraph
ownParagraph.AppendField('MyFieldName', spiredoc.FieldType.FieldMergeField);
//Put these objects back into the paragraph one by one
for (let obj of tempList) {
  ownParagraph.ChildObjects.Add(obj);
}
```

---

# spire.doc javascript text replacement
## replace text with table in word document
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

//Return TextSection by finding the key text string "Christmas Day, December 25".
let section = doc.Sections.get_Item(0);
let selection = doc.FindString('Christmas Day, December 25', true, true);

//Return TextRange from TextSection, then get OwnerParagraph through TextRange.
let range = selection.GetAsOneRange();
let paragraph = range.OwnerParagraph;

//Return the zero-based index of the specified paragraph.
let body = paragraph.OwnerTextBody;
let index = body.ChildObjects.IndexOf(paragraph);

//Create a new table.
let table = section.AddTable(true);
table.ResetCells(3, 3);

//Remove the paragraph and insert table into the collection at the specified index.
body.ChildObjects.Remove(paragraph);
body.ChildObjects.Insert(index, table);
```

---

# Spire.Doc JavaScript Replace Text
## Replace text in a Word document with another Word document
```javascript
// Load a template document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName1);

// Load another document to replace text
let replaceDoc = wasmModule.Document.Create();
replaceDoc.LoadFromFile(inputFileName2);

// Replace specified text with the other document
doc.Replace({
  matchString: 'Document1',
  matchDoc: replaceDoc,
  caseSensitive: false,
  wholeWord: true,
});
```

---

# Word Document Find and Replace with HTML
## Replace text placeholders with HTML content in a Word document
```javascript
// Collect the objects which is used to replace text
let replacement = [];

// Create a temporary section
let tempSection = document.AddSection();

// Add a paragraph to append html
let par = tempSection.AddParagraph();

// Append the HTML content to the paragraph
par.AppendHTML(htmlString);

// Get the objects in temporary section
for (let i = 0; i < tempSection.Body.ChildObjects.Count; i++) {
  let docObj = tempSection.Body.ChildObjects.get(i);
  replacement.push(docObj);
}

//Find all text which will be replaced.
let selections = document.FindAllString('[#placeholder]', false, true);
let locations = [];
for (let selection of selections) {
  // Get the range of the current selection and create a new TextRangeLocation object with it
  locations.push(new TextRangeLocation(selection.GetAsOneRange()));
}
locations.sort();

for (let location of locations) {
  //replace the text with HTML
  ReplaceWithHTML(location, replacement);
}

//remove the temp section
document.Sections.Remove(tempSection);

function ReplaceWithHTML(location, replacement) {
  let textRange = location.Text;
  
  //textRange index
  let index = location.Index;
  
  //get owener paragraph
  let paragraph = location.Owner;
  
  //get owner text body
  let sectionBody = paragraph.OwnerTextBody;
  
  //get the index of paragraph in section
  let paragraphIndex = sectionBody.ChildObjects.IndexOf(paragraph);
  
  let replacementIndex = -1;
  if (index === 0) {
    //remove the first child object
    paragraph.ChildObjects.RemoveAt(0);
    
    replacementIndex = sectionBody.ChildObjects.IndexOf(paragraph);
  } else if (index == paragraph.ChildObjects.Count - 1) {
    paragraph.ChildObjects.RemoveAt(index);
    replacementIndex = paragraphIndex + 1;
  } else {
    //split owner paragraph
    let paragraph1 = paragraph.Clone();
    while (paragraph.ChildObjects.Count > index) {
      paragraph.ChildObjects.RemoveAt(index);
    }
    let i = 0;
    let count = index + 1;
    while (i < count) {
      paragraph1.ChildObjects.RemoveAt(0);
      i += 1;
    }
    sectionBody.ChildObjects.Insert(paragraphIndex + 1, paragraph1);
    
    replacementIndex = paragraphIndex + 1;
  }
  
  //insert replacement
  for (let i = 0; i <= replacement.length - 1; i++) {
    sectionBody.ChildObjects.Insert(replacementIndex + i, replacement[i].Clone());
  }
}

function TextRangeLocation(TextRange) {
  this.Text = TextRange;
  this.Owner = this.Text.OwnerParagraph;
  this.Index = this.Owner.ChildObjects.IndexOf(this.Text);
  this.CompareTo = function (other) {
    return -(this.Index - other.Index);
  };
}
```

---

# spire.doc javascript find and replace
## replace text with image in word document
```javascript
//Find the string "E-iceblue" in the document
let selections = document.FindAllString('E-iceblue', true, true);
let index = 0;
let range = null;

//Remove the text and replace it with Image
for (let i = 0; i < selections.length; i++) {
  // Create a new DocPicture object and load the defined image into it
  let pic = wasmModule.DocPicture.Create(document);
  pic.LoadImage(pngName);
  let selection = selections[i];
  // Get the current range of text being processed
  range = selection.GetAsOneRange();
  // Get the current index of the TextRange within its owner paragraph's ChildObjects collection
  index = range.OwnerParagraph.ChildObjects.IndexOf(range);
  // Insert the image into the owner paragraph's ChildObjects collection at the position of the TextRange
  range.OwnerParagraph.ChildObjects.Insert(index, pic);
  // Remove the TextRange from its owner paragraph's ChildObjects collection
  range.OwnerParagraph.ChildObjects.Remove(range);
}
```

---

# Spire.Doc JavaScript Text Replacement
## Replace text in a Word document
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Replace text
doc.Replace({ matchString: 'word', newValue: 'ReplacedText', caseSensitive: false, wholeWord: true });

// Define the output file name
const outputFileName = 'ReplaceWithText.docx';

// Save the document to the specified path
doc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });
```

---

# spire.doc javascript content extraction
## extract content between paragraphs
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

//Create a destination document
let destinationDoc = wasmModule.Document.Create();

//Add a section
let section = destinationDoc.AddSection();

//Extract content between the first paragraph to the third paragraph
ExtractBetweenParagraphs(doc, destinationDoc, 1, 3);

function ExtractBetweenParagraphs(sourceDocument, destinationDocument, startPara, endPara) {
  //Extract the content
  for (let i = startPara - 1; i < endPara; i++) {
    //Clone the ChildObjects of source document
    let doobj = sourceDocument.Sections.get_Item(0).Body.ChildObjects.get_Item(i).Clone();

    //Add to destination document
    destinationDocument.Sections.get_Item(0).Body.ChildObjects.Add(doobj);
  }
}
```

---

# Document Content Extraction
## Extract content between paragraph styles
```javascript
// Create a new document
const sourceDocument = wasmModule.Document.Create();

// Load a document from the virtual file system
sourceDocument.LoadFromFile(inputFileName);

// Create a destination document
let destinationDoc = wasmModule.Document.Create();

// Add a section
let section = destinationDoc.AddSection();
let stylename1 = '1';
let stylename2 = '2';
// Extract content between the first paragraph to the third paragraph
let startindex = 0;
let endindex = 0;
// Travel the sections of source document
for (let i = 0; i < sourceDocument.Sections.Count; i++) {
  let section1 = sourceDocument.Sections.get(i);
  // Travel the paragraphs
  for (let j = 0; j < section1.Paragraphs.Count; j++) {
    let paragraph = section1.Paragraphs.get_Item(j);
    // Judge paragraph style1
    if (paragraph.StyleName === stylename1) {
      // Get the paragraph index
      startindex = section1.Body.Paragraphs.IndexOf(paragraph);
    }
    // Judge paragraph style2
    if (paragraph.StyleName === stylename2) {
      // Get the paragraph index
      endindex = section1.Body.Paragraphs.IndexOf(paragraph);
    }
  }
  // Extract the content
  for (let i = startindex + 1; i < endindex; i++) {
    // Clone the ChildObjects of source document
    let doobj = sourceDocument.Sections.get_Item(0).Body.ChildObjects.get_Item(i).Clone();

    // Add to destination document
    destinationDoc.Sections.get_Item(0).Body.ChildObjects.Add(doobj);
  }
}
```

---

# Spire.Doc JavaScript Paragraph Extraction
## Extract paragraphs from document based on style name
```javascript
let styleName1 = 'Heading1';
let style1Text = '';
style1Text += 'The following is the content of the paragraph with the style name ' + styleName1 + ': ' + '\n';
// Extract paragraph based on style
for (let i = 0; i < doc.Sections.Count; i++) {
  let section = doc.Sections.get_Item(i);
  // Travel through the paragraphs
  for (let j = 0; j < section.Paragraphs.Count; j++) {
    let paragraph = section.Paragraphs.get_Item(j);
    if (paragraph.StyleName != null && paragraph.StyleName === styleName1) {
      style1Text += paragraph.Text;
    }
  }
}
```

---

# Extract Content from Bookmark in Document
## This code demonstrates how to extract content from a bookmark in a Word document and add it to a destination document.
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document
doc.LoadFromFile(inputFileName);

//Create a destination document
let destinationDoc = wasmModule.Document.Create();

//Add a section for destination document
let section = destinationDoc.AddSection();

//Add a paragraph for destination document
let paragraph = section.AddParagraph();

//Locate the bookmark in source document
let navigator = wasmModule.BookmarksNavigator.Create(doc);

//Find bookmark by name
navigator.MoveToBookmark({
  bookmarkName: 'Test',
  isStart: true,
  isAfter: true,
});

//get text body part
let textBodyPart = navigator.GetBookmarkContent();

//Create a TextRange type list
let list = [];

//Traverse the items of text body
for (let i = 0; i < textBodyPart.BodyItems.Count; i++) {
  let item = textBodyPart.BodyItems.get(i);

  for (let j = 0; j < item.ChildObjects.Count; j++) {
    let childObject = item.ChildObjects.get(j);

    //Add it into list
    let range = childObject;
    list.push(range);
  }
}

//Add the extract content to destinationDoc document
for (let m = 0; m < list.length; m++) {
  paragraph.Items.Add(list[m].Clone());
}
```

---

# Spire.Doc JavaScript Comment Content Extraction
## Extract content from comment range in a Word document

```javascript
//Create a source document
let sourceDoc = wasmModule.Document.Create();

//Create a destination document
let destinationDoc = wasmModule.Document.Create();

//Add section for destination document
let destinationSec = destinationDoc.AddSection();

//Get the first comment
let comment = sourceDoc.Comments.get_Item(0);

//Get the paragraph of obtained comment
let para = comment.OwnerParagraph;

//Get index of the CommentMarkStart
let startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart);

//Get index of the CommentMarkEnd
let endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd);

//Traverse paragraph ChildObjects
for (let i = startIndex; i <= endIndex; i++) {
  //Clone the ChildObjects of source document
  let doobj = para.ChildObjects.get(i).Clone();

  //Add to destination document
  destinationSec.AddParagraph().ChildObjects.Add(doobj);
}
```

---

# Spire.Doc JavaScript Content Extraction
## Extract content between a paragraph and a table
```javascript
function ExtractByTable(sourceDocument, destinationDocument, startPara, tableNo) {
  //Get the table from the source document
  let table = sourceDocument.Sections.get(0).Tables.get_Item(tableNo - 1);

  //Get the table index
  let index = sourceDocument.Sections.get_Item(0).Body.ChildObjects.IndexOf(table);
  for (let i = startPara - 1; i <= index; i++) {
    //Clone the ChildObjects of source document
    let doobj = sourceDocument.Sections.get(0).Body.ChildObjects.get(i).Clone();

    //Add to destination document
    destinationDocument.Sections.get(0).Body.ChildObjects.Add(doobj);
  }
}
```

---

# Extract Content from Form Field
## Demonstrates how to extract content starting from a form field in a Word document
```javascript
// Create the source document
let sourceDocument = wasmModule.Document.Create();

// Load the source document from disk
sourceDocument.LoadFromFile(inputFileName);

// Create a destination document
let destinationDoc = wasmModule.Document.Create();

// Add a section
let section = destinationDoc.AddSection();

// Define a variables
let index = 0;
let formFields = sourceDocument.Sections.get(0).Body.FormFields;

// Traverse FormFields
for (let i = 0; i < formFields.Count; i++) {
  let field = formFields.get_Item(i);
  
  // Find FieldFormTextInput type field
  if (field.Type == wasmModule.FieldType.FieldFormTextInput) {
    // Get the paragraph
    let paragraph = field.OwnerParagraph;

    // Get the index
    index = sourceDocument.Sections.get(0).Body.ChildObjects.IndexOf(paragraph);
    break;
  }
}

// Extract the content
for (let i = index; i < index + 3; i++) {
  // Clone the ChildObjects of source document
  let doobj = sourceDocument.Sections.get(0).Body.ChildObjects.get(i).Clone();

  // Add to destination document
  section.Body.ChildObjects.Add(doobj);
}
```

---

# spire.doc javascript sections
## add and delete sections in a Word document
```javascript
function AddSection(doc) {
  // Add a section
  doc.AddSection();
}

function DeleteSection(doc) {
  // Delete the last section
  doc.Sections.RemoveAt(doc.Sections.Count - 1);
}
```

---

# Spire.Doc JavaScript Section Cloning
## Clone sections from one Word document to another
```javascript
//Load source file
let srcDoc = wasmModule.Document.Create();

//Create destination file
let desDoc = wasmModule.Document.Create();

let cloneSection = null;
for (let i = 0; i < srcDoc.Sections.Count; i++) {
  //Clone section
  cloneSection = srcDoc.Sections.get(i).Clone();
  //Add the cloneSection in destination file
  desDoc.Sections.Add(cloneSection);
}
```

---

# Clone Section Content in Word Document
## This code demonstrates how to clone content from one section to another in a Word document
```javascript
//Load the Word document from disk
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Get the first section
let sec1 = doc.Sections.get_Item(0);
//Get the second section
let sec2 = doc.Sections.get_Item(1);

//Loop through the contents of sec1
for (let i = 0; i < sec1.Body.ChildObjects.Count; i++) {
  //Clone the contents to sec2
  sec2.Body.ChildObjects.Add(sec1.Body.ChildObjects.get(i).Clone());
}
```

---

# Spire.Doc JavaScript Page Setup
## Modify page setup properties of document sections
```javascript
// Load Word document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

// Loop through all sections
for (let i = 0; i < doc.Sections.Count; i++) {
  let section = doc.Sections.get(i);
  // Modify the margins
  section.PageSetup.Margins = wasmModule.MarginsF.Create(100, 80, 100, 80);
  // Modify the page size
  section.PageSetup.PageSize = wasmModule.PageSize.Letter;
}

// Modify only one section (first section)
let section0 = doc.Sections.get_Item(0);
section0.PageSetup.Margins = wasmModule.MarginsF.Create(100, 80, 100, 80);
section0.PageSetup.FooterDistance = 35.4;
section0.PageSetup.HeaderDistance = 34.4;
```

---

# Word Document Section Content Removal
## This code demonstrates how to remove header, body, and footer content from all sections in a Word document
```javascript
// Loop through all sections
for (let i = 0; i < doc.Sections.Count; i++) {
  let section = doc.Sections.get(0);
  // Remove header content
  section.HeadersFooters.Header.ChildObjects.Clear();
  // Remove body content
  section.Body.ChildObjects.Clear();
  // Remove footer content
  section.HeadersFooters.Footer.ChildObjects.Clear();
}
```

---

# spire.doc javascript paragraph
## add tab stops to word paragraphs
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Add a section.
let section = document.AddSection();

//Add paragraph 1.
let paragraph1 = section.AddParagraph();

//Add tab and set its position (in points).
let tab = paragraph1.Format.Tabs.AddTab({position: 28});

//Set tab alignment.
tab.Justification = wasmModule.TabJustification.Left;

//Move to next tab and append text.
paragraph1.AppendText('\tWashing Machine');

//Add another tab and set its position (in points).
tab = paragraph1.Format.Tabs.AddTab({position: 280});

//Set tab alignment.
tab.Justification = wasmModule.TabJustification.Left;

//Specify tab leader type.
tab.TabLeader = wasmModule.TabLeader.Dotted;

//Move to next tab and append text.
paragraph1.AppendText('\t$650');

//Add paragraph 2.
let paragraph2 = section.AddParagraph();

//Add tab and set its position (in points).
tab = paragraph2.Format.Tabs.AddTab({position: 28});

//Set tab alignment.
tab.Justification = wasmModule.TabJustification.Left;

//Move to next tab and append text.
paragraph2.AppendText('\tRefrigerator');

//Add another tab and set its position (in points).
tab = paragraph2.Format.Tabs.AddTab({position: 280});

//Set tab alignment.
tab.Justification = wasmModule.TabJustification.Left;

//Specify tab leader type.
tab.TabLeader = wasmModule.TabLeader.NoLeader;

//Move to next tab and append text.
paragraph2.AppendText('\t$800');
```

---

# Spire.Doc JavaScript Paragraph Formatting
## Allow Latin text wrap in middle of a word
```javascript
let para = document.Sections.get(0).Paragraphs.get_Item(0);
//Allow Latin text to wrap in the middle of a word
para.Format.WordWrap = false;
```

---

# spire.doc javascript paragraph operations
## copy paragraphs between word documents
```javascript
//Create Word document1.
let document1 = wasmModule.Document.Create();

//Load the file from disk.
document1.LoadFromFile(inputFileName);

//Create a new document.
let document2 = wasmModule.Document.Create();

//Get paragraph 1 and paragraph 2 in document1.
let s = document1.Sections.get(0);
let p1 = s.Paragraphs.get_Item(0);
let p2 = s.Paragraphs.get_Item(1);

//Copy p1 and p2 to document2.
let s2 = document2.AddSection();
let NewPara1 = p1.Clone();
s2.Paragraphs.Add(NewPara1);

let NewPara2 = p2.Clone();
s2.Paragraphs.Add(NewPara2);

//Add watermark.
let WM = wasmModule.PictureWatermark.Create();
// Set the Picture property of WM to an image
WM.SetPicture(inputFileName2);
// Set the Watermark property of document2 to WM
document2.Watermark = WM;
```

---

# Word Document Catalogue Formation
## Create a catalogue from Word headings with numbered list styles
```javascript
//Create Word document
let document = wasmModule.Document.Create();

//Add a new section
let section = document.AddSection();
let paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();

//Add Heading 1
paragraph = section.AddParagraph();
paragraph.AppendText("Heading1");
paragraph.ApplyStyle({
  builtinStyle: wasmModule.BuiltinStyle.Heading1,
});
paragraph.ListFormat.ApplyNumberedStyle();

//Add Heading 2
paragraph = section.AddParagraph();
paragraph.AppendText("Heading2");
paragraph.ApplyStyle({
  builtinStyle: wasmModule.BuiltinStyle.Heading2,
});

//List style for Headings 2
let listSty2 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Numbered);
for (let i = 0; i < listSty2.Levels.Count; i++) {
  let listLev = listSty2.Levels.get_Item(i);
  listLev.UsePrevLevelPattern = true;
  listLev.NumberPrefix = '1.';
}
listSty2.Name = 'MyStyle2';
document.ListStyles.Add(listSty2);
paragraph.ListFormat.ApplyStyle(listSty2.Name);

//Add list style 3
let listSty3 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Numbered);
for (let i = 0; i < listSty3.Levels.Count; i++) {
  let listLev = listSty3.Levels.get_Item(i);
  listLev.UsePrevLevelPattern = true;
  listLev.NumberPrefix = '1.1.';
}
listSty3.Name = 'MyStyle3';
document.ListStyles.Add(listSty3);

//Add Heading 3
for (let i = 0; i < 4; i++) {
  paragraph = section.AddParagraph();
  
  //Append text
  paragraph.AppendText("Heading3");
  
  //Apply list style 3 for Heading 3
  paragraph.ApplyStyle({
    builtinStyle: wasmModule.BuiltinStyle.Heading3,
  });
  paragraph.ListFormat.ApplyStyle(listSty3.Name);
}
```

---

# spire.doc javascript paragraph
## get paragraphs by style name
```javascript
//Create Word document
let document = wasmModule.Document.Create();

//Load the file from disk
document.LoadFromFile(inputFileName);

let content = [];
content.push('Get paragraphs by style name "Heading1": ' + '\n');

//Get paragraphs by style name
for (let i = 0; i < document.Sections.Count; i++) {
  let section = document.Sections.get_Item(i);
  for (let j = 0; j < section.Paragraphs.Count; j++) {
    let paragraph = section.Paragraphs.get_Item(j);
    if (paragraph.StyleName == 'Heading1') {
      content.push(paragraph.Text);
    }
  }
}
```

---

# Spire.Doc JavaScript Paragraph Revisions
## Get revision details of paragraphs in a document
```javascript
let document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

let builder = [];

//loop paragraph
for (let i = 0; i < document.Sections.Count; i++) {
  let section = document.Sections.get_Item(0);
  for (let j = 0; j < section.Paragraphs.Count; j++) {
    let paragraph = section.Paragraphs.get_Item(j);
    // Check if the Paragraph is a deleted revision.
    if (paragraph.IsDeleteRevision) {
      // Append information about the deleted revision to the builder.
      builder.push('The section' + document.GetIndex(section) + 'paragraph' + section.GetIndex(paragraph) + 'has been changed (deleted).' + '\n');
      builder.push('Author: ' + paragraph.DeleteRevision.Author + '\n');
      builder.push('DateTime: ' + paragraph.DeleteRevision.DateTime.ToString() + '\n');
      builder.push('Type: ' + paragraph.DeleteRevision.Type + '\n');
      builder.push('' + '\n');
    }
    // Check if the Paragraph is an inserted revision.
    else if (paragraph.IsInsertRevision) {
      // Append information about the inserted revision to the builder.
      builder.push('The section' + document.GetIndex(section) + 'paragraph' + section.GetIndex(paragraph) + 'has been changed (inserted).' + '\n');
      builder.push('Author: ' + paragraph.InsertRevision.Author + '\n');
      builder.push('DateTime: ' + paragraph.InsertRevision.DateTime.ToString() + '\n');
      builder.push('Type: ' + paragraph.InsertRevision.Type + '\n');
      builder.push('' + '\n');
    }
    // Iterate over the child DocumentObjects in the Paragraph.
    else {
      for (let i = 0; i < paragraph.ChildObjects.Count; i++) {
        let obj = paragraph.ChildObjects.get(i);
        // Check if the child DocumentObject is a TextRange.
        if (obj.DocumentObjectType == wasmModule.DocumentObjectType.TextRange) {
          let textRange = obj;
          {
            // Check if the TextRange is a deleted revision.
            if (textRange.IsDeleteRevision) {
              builder.push(
                'The section' + document.GetIndex(section) + 'paragraph' + section.GetIndex(paragraph) + 'textrange' + paragraph.GetIndex(textRange) + 'has been changed (deleted).' + '\n'
              );
              builder.push('Author: ' + textRange.DeleteRevision.Author + '\n');
              builder.push('DateTime: ' + textRange.DeleteRevision.DateTime.ToString() + '\n');
              builder.push('Type: ' + textRange.DeleteRevision.Type + '\n');
              builder.push('Change Text: ' + textRange.Text + '\n');
              builder.push('' + '\n');
            }
            // Check if the TextRange is an inserted revision.
            else if (textRange.IsInsertRevision) {
              builder.push(
                'The section' + document.GetIndex(section) + 'paragraph' + section.GetIndex(paragraph) + 'textrange' + paragraph.GetIndex(textRange) + 'has been changed (deleted).' + '\n'
              );
              builder.push('Author: ' + textRange.InsertRevision.Author + '\n');
              builder.push('DateTime: ' + textRange.InsertRevision.DateTime.ToString() + '\n');
              builder.push('Type: ' + textRange.InsertRevision.Type + '\n');
              builder.push('Change Text: ' + textRange.Text + '\n');
              builder.push('' + '\n');
            }
          }
        }
      }
    }
  }
}
```

---

# spire.doc javascript paragraph
## hide paragraph in word document
```javascript
//Get the first section and the first paragraph from the word document.
let sec = document.Sections.get(0);
let para = sec.Paragraphs.get_Item(0);

//Loop through the textranges and set CharacterFormat.Hidden property as true to hide the texts.
for (let i = 0; i < para.ChildObjects.Count; i++) {
  let obj = para.ChildObjects.get(i);
  if (obj instanceof wasmModule.TextRange) {
    let range = obj;
    range.CharacterFormat.Hidden = true;
  }
}
```

---

# spire.doc javascript RTF insertion
## insert RTF string into Word document
```javascript
//Create Word document
let document = wasmModule.Document.Create();

//Add a new section
let section = document.AddSection();

//Add a paragraph to the section
let para = section.AddParagraph();

//Declare a String variable to store the Rtf string
let rtfString = '{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 hakuyoxingshu7000;}}\\f0\\fs28 Hello, World}';

//Append Rtf string to paragraph
para.AppendRTF(rtfString);
```

---

# Word Document Pagination Management
## Set page break before a paragraph in Word document
```javascript
//Get the first section and the paragraph we want to manage the pagination.
let sec = document.Sections.get(0);
let para = sec.Paragraphs.get_Item(4);

//Set the pagination format as Format.PageBreakBefore for the checked paragraph.
para.Format.PageBreakBefore = true;
```

---

# Remove All Paragraphs in Word Document
## This code demonstrates how to remove all paragraphs from every section in a Word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file from disk.
document.LoadFromFile(inputFileName);

//Remove paragraphs from every section in the document
for (let i = 0; i < document.Sections.Count; i++) {
  document.Sections.get(i).Paragraphs.Clear();
}
```

---

# spire.doc javascript paragraph
## remove empty lines from word document
```javascript
// Traverse every section on the word document and remove the null and empty paragraphs
for (let i = 0; i < document.Sections.Count; i++) {
  let section = document.Sections.get_Item(i);
  for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
    if (section.Body.ChildObjects.get(j).DocumentObjectType === wasmModule.DocumentObjectType.Paragraph) {
      let str = section.Body.ChildObjects.get(j).Text.trim();
      if (str.length === 0) {
        section.Body.ChildObjects.Remove(section.Body.ChildObjects.get(j));
        j--;
      }
    }
  }
}
```

---

# Spire.Doc for JavaScript
## Remove a specific paragraph from a Word document
```javascript
// Create Word document
let document = wasmModule.Document.Create();

// Load the file from disk
document.LoadFromFile(inputFileName);

// Remove the first paragraph from the first section of the document
document.Sections.get(0).Paragraphs.RemoveAt(0);
```

---

# spire.doc javascript frame position
## set frame position in word document
```javascript
//Get a paragraph
let paragraph = document.Sections.get(0).Paragraphs.get_Item();

//Set the Frame's position
if (paragraph.Format.IsFrame) {
  paragraph.Format.Frame.SetHorizontalPosition(150);
  paragraph.Format.Frame.SetVerticalPosition(150);
}
```

---

# Spire.Doc JavaScript Paragraph Shading
## Set background color for paragraphs in Word documents
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file from disk.
document.LoadFromFile(inputFileName);
//Get a paragraph.
let paragaph = document.Sections.get(0).Paragraphs.get_Item(0);

//Set background color for the paragraph.
paragaph.Format.BackColor = wasmModule.Color.get_Yellow();

//Set background color for the selected text of paragraph.
paragaph = document.Sections.get(0).Paragraphs.get_Item(2);
let selection = paragaph.Find({
  given: 'Christmas',
  caseSensitive: true,
  wholeWord: false,
});
let range = selection.GetAsOneRange();
range.CharacterFormat.TextBackgroundColor = wasmModule.Color.get_Yellow();
```

---

# Spire.Doc JavaScript Paragraph Formatting
## Set space between Asian and Latin text
```javascript
let para = document.Sections.get(0).Paragraphs.get_Item(0);

//Set whether to automatically adjust space between Asian text and Latin text
para.Format.AutoSpaceDE = false;
//Set whether to automatically adjust space between Asian text and numbers
para.Format.AutoSpaceDN = true;
```

---

# spire.doc javascript paragraph
## set paragraph before and after spacing
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Add the text strings to the paragraph and set the style.
let para = wasmModule.Paragraph.Create(document);
let str = 'This is an inserted paragraph.';
let textRange1 = para.AppendText(str);
textRange1.CharacterFormat.TextColor = wasmModule.Color.get_Blue();
textRange1.CharacterFormat.FontSize = 15;

//set the spacing before and after.
para.Format.BeforeAutoSpacing = false;
para.Format.BeforeSpacing = 10;
para.Format.AfterAutoSpacing = false;
para.Format.AfterSpacing = 10;

//insert the added paragraph to the word document.
document.Sections.get(0).Paragraphs.Insert(1, para);
```

---

# spire.doc javascript text emphasis
## apply emphasis mark to text in word document
```javascript
// Create a new document and load from file
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

// Find text to emphasize
let textSelections = doc.FindAllString('Spire.Doc for JavaScript', false, true);

// Set emphasis mark to the found text
for (let i = 0; i < textSelections.length; i++) {
  let selection = textSelections[i];
  // Get the found text range as a single range and apply an emphasis mark (dot) to its character format
  selection.GetAsOneRange().CharacterFormat.EmphasisMark = wasmModule.Emphasis.Dot;
}
```

---

# Word Document Text Case Conversion
## Change text case to capital letters and small caps in a Word document
```javascript
// Create a new document and load from file
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);
let textRange;

// Get the first paragraph and set its CharacterFormat to AllCaps
let para1 = doc.Sections.get(0).Paragraphs.get_Item(1);

for (let i = 0; i < para1.ChildObjects.Count; i++) {
  let obj = para1.ChildObjects.get(i);
  if (obj instanceof wasmModule.TextRange) {
    textRange = obj;
    textRange.CharacterFormat.AllCaps = true;
  }
}

// Get the third paragraph and set its CharacterFormat to IsSmallCaps
let para2 = doc.Sections.get(0).Paragraphs.get_Item(3);
for (let j = 0; j < para2.ChildObjects.Count; j++) {
  let obj = para2.ChildObjects.get(j);
  if (obj instanceof wasmModule.TextRange) {
    obj.CharacterFormat.IsSmallCaps = true;
  }
}
```

---

# Spire.Doc JavaScript Barcode
## Create barcode in Word document
```javascript
// Create a document
let doc = wasmModule.Document.Create();

// Add a paragraph
let p = doc.AddSection().AddParagraph();

// Add barcode and set its format
let txtRang = p.AppendText('H63TWX11072');
// Set barcode font name, note you need to install the barcode font on your system at first
txtRang.CharacterFormat.FontName = 'C39HrP60DlTt';
txtRang.CharacterFormat.FontSize = 80;
txtRang.CharacterFormat.TextColor = wasmModule.Color.get_SeaGreen();
```

---

# Extract Text from Word Document
## This example demonstrates how to extract text from a Word document using Spire.Doc for JavaScript
```javascript
//Load the document from disk.
let document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

//get text from document
let text = document.GetText();
```

---

# Word Document Text Insertion
## Insert new text after searched text and highlight it
```javascript
//Load Document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Find all the text string "Word" from the sample document
let selections = doc.FindAllString('Word', true, true);
let index = 0;

//Defines text range
let range;

//Insert new text string (New) after the searched text string
for (let i = 0; i < selections.length; i++) {
  let selection = selections[i];
  range = selection.GetAsOneRange();
  let newrange = wasmModule.TextRange.Create(doc);
  newrange.Text = '(New text)';
  index = range.OwnerParagraph.ChildObjects.IndexOf(range);
  range.OwnerParagraph.ChildObjects.Insert(index + 1, newrange);
}

//Find and highlight the newly added text string New
let text = doc.FindAllString('New text', true, true);
for (let i = 0; i < text.length; i++) {
  let seletion = text[i];
  seletion.GetAsOneRange().CharacterFormat.HighlightColor = wasmModule.Color.get_Yellow();
}
```

---

# Spire.Doc JavaScript Symbol Insertion
## Insert symbols in Word document using Unicode characters
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Add a section.
let section = document.AddSection();

//Add a paragraph.
let paragraph = section.AddParagraph();

//Use unicode characters to create symbol Ä.
let tr = paragraph.AppendText('\u00c4'.toString());

//Set the color of symbol Ä.
tr.CharacterFormat.TextColor = wasmModule.Color.get_Red();

//Add symbol Ë.
paragraph.AppendText('\u00cb'.toString());
```

---

# Spire.Doc JavaScript Text Loading
## Load text with UTF-7 encoding
```javascript
let doc = wasmModule.Document.Create();
doc.LoadText({ fileName: inputFileName, encoding: wasmModule.Encoding.get_UTF7() });

// Save the document to the specified path
doc.SaveToFile({
  fileName: outputFileName,
  fileFormat: wasmModule.FileFormat.Docx2013,
});

// Clean up resources
doc.Dispose();
```

---

# Setting Superscript and Subscript in Word Document
## Demonstrates how to apply superscript and subscript formatting to text ranges in a Word document
```javascript
// Create word document
let document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

let paragraph = section.AddParagraph();
paragraph.AppendText('E = mc');
let range1 = paragraph.AppendText('2');

// Set superscript
range1.CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SuperScript;
// Insert a line break in the paragraph.
paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
paragraph.AppendText('F');
let range2 = paragraph.AppendText('n');

// Set subscript
range2.CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SubScript;
// Append the text " = Fn-1 + Fn-2" with specific subscripts to the paragraph.
paragraph.AppendText(' = F');
paragraph.AppendText('n-1').CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SubScript;
paragraph.AppendText(' + F');
paragraph.AppendText('n-2').CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SubScript;

// Set font size
for (let i = 0; i < paragraph.Items.Count; i++) {
  let range = paragraph.Items.get_Item(i);
  if (range instanceof wasmModule.TextRange) {
    range.CharacterFormat.FontSize = 36;
  }
}
```

---

# Setting Text Direction in Word Document
## Demonstrates how to set text direction in Word document sections and table cells
```javascript
//Create a new document
let doc = wasmModule.Document.Create();

//Add the first section
let section1 = doc.AddSection();
//Set text direction for all text in a section
section1.TextDirection = wasmModule.TextDirection.RightToLeft;

//Set Font Style and Size
let style = wasmModule.ParagraphStyle.Create(doc);
style.Name = 'FontStyle';
style.CharacterFormat.FontName = 'Arial';
style.CharacterFormat.FontSize = 15;

doc.Styles.Add(style);

//Add two paragraphs and apply the font style
let p = section1.AddParagraph();
p.AppendText('Only Spire.Doc, no Microsoft Office automation');
p.ApplyStyle({ styleName: style.Name });
p = section1.AddParagraph();
p.AppendText('Convert file documents with high quality');
p.ApplyStyle({ styleName: style.Name });

//Set text direction for a part of text
//Add the second section
let section2 = doc.AddSection();
//Add a table
let table = section2.AddTable();
table.ResetCells(1, 1);
let cell = table.Rows.get(0).Cells.get(0);
table.Rows.get(0).Height = 150;
table.Rows.get(0).Cells.get(0).SetCellWidth(10, wasmModule.CellWidthType.Point);
//Set vertical text direction of table
cell.CellFormat.TextDirection = wasmModule.TextDirection.RightToLeftRotated;
cell.AddParagraph().AppendText('This is vertical style');
//Add a paragraph and set horizontal text direction
p = section2.AddParagraph();
p.AppendText('This is horizontal style');
p.ApplyStyle(style.Name);
```

---

# Word Document Text Splitting
## Split text into columns in a Word document
```javascript
//Create a new document and load from file
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Add a column to the first section and set width and spacing
doc.Sections.get_Item(0).AddColumn(100, 20);
//Add a line between the two columns
doc.Sections.get_Item(0).PageSetup.ColumnsLineBetween = true;
```

---

# Alter Language Dictionary in Word Document
## Change language settings for text in a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Add new section and paragraph to the document
let sec = document.AddSection();
let para = sec.AddParagraph();

// Add a textRange for the paragraph and append some Peru Spanish words
let txtRange = para.AppendText("corrige según diccionario en inglés");
txtRange.CharacterFormat.LocaleIdASCII = 10250;

// Save the document to the specified path
document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

// Clean up resources
document.Dispose();
```

---

# Document Format Detection
## Check file format of Word documents
```javascript
// Create a new document
const document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

//Get file format
let ff = document.DetectedFormatType;
let fileFormat = "The file format is ";

//Check the format info
switch (ff) {
  case wasmModule.FileFormat.Doc:
    fileFormat += "Microsoft Word 97-2003 document.";
    break;
  case wasmModule.FileFormat.Dot:
    fileFormat += "Microsoft Word 97-2003 template.";
    break;
  case wasmModule.FileFormat.Docx:
    fileFormat += "Office Open XML WordprocessingML Macro-Free Document.";
    break;
  case wasmModule.FileFormat.Docm:
    fileFormat += "Office Open XML WordprocessingML Macro-Enabled Document.";
    break;
  case wasmModule.FileFormat.Dotx:
    fileFormat += "Office Open XML WordprocessingML Macro-Free Template.";
    break;
  case wasmModule.FileFormat.Dotm:
    fileFormat += "Office Open XML WordprocessingML Macro-Enabled Template.";
    break;
  case wasmModule.FileFormat.Rtf:
    fileFormat += "RTF format.";
    break;
  case wasmModule.FileFormat.WordML:
    fileFormat += "Microsoft Word 2003 WordprocessingML format.";
    break;
  case wasmModule.FileFormat.Html:
    fileFormat += "HTML format.";
    break;
  case wasmModule.FileFormat.WordXml:
    fileFormat += "Microsoft Word xml format for word 2007-2013.";
    break;
  case wasmModule.FileFormat.Odt:
    fileFormat += "OpenDocument Text.";
    break;
  case wasmModule.FileFormat.Ott:
    fileFormat += "OpenDocument Text Template.";
    break;
  case wasmModule.FileFormat.DocPre97:
    fileFormat += "Microsoft Word 6 or Word 95 format.";
    break;
  default:
    fileFormat += "Unknown format.";
    break;
}
```

---

# Spire.Doc JavaScript Document Comparison
## Compare two Word documents and highlight differences
```javascript
//Load the first document
let doc1 = wasmModule.Document.Create();
doc1.LoadFromFile(inputFileName1);

//Load the second document
let doc2 = wasmModule.Document.Create();
doc2.LoadFromFile(inputFileName2);

//Compare the two documents
doc1.Compare(doc2, "E-iceblue");

// Save the document to the specified path
doc1.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });
```

---

# Spire.Doc Document Comparison
## Compare two Word documents with specific options
```javascript
// Load the first document
let doc1 = wasmModule.Document.Create();
doc1.LoadFromFile(inputFileName1);

// Load the second document
let doc2 = wasmModule.Document.Create();
doc2.LoadFromFile(inputFileName2);

// Create compareOptions
let compareOptions = wasmModule.CompareOptions.Create();
compareOptions.IgnoreFormatting = true;

// Compare the two documents
doc1.Compare({
  document: doc2, 
  author: "E-iceblue", 
  dateTime: wasmModule.DateTime.get_Now(), 
  options: compareOptions
});

// Save the document to the specified path
doc1.SaveToFile({ 
  fileName: outputFileName, 
  fileFormat: wasmModule.FileFormat.Docx2013 
});
```

---

# Spire.Doc JavaScript Word Count
## Count characters and words in a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

// Count the number of words.
let content = [];
content.push("CharCount: " + document.BuiltinDocumentProperties.CharCount + "\n");
content.push("CharCountWithSpace: " + document.BuiltinDocumentProperties.CharCountWithSpace + "\n");
content.push("WordCount: " + document.BuiltinDocumentProperties.WordCount + "\n");
```

---

# Spire.Doc JavaScript Document Property
## Set Word document properties
```javascript
wasmModule = window.wasmModule;
if (wasmModule) {
  // Create a new document
  const document = wasmModule.Document.Create();

  // Set document properties
  document.BuiltinDocumentProperties.Title = "Document Demo Document";
  document.BuiltinDocumentProperties.Subject = "demo";
  document.BuiltinDocumentProperties.Author = "James";
  document.BuiltinDocumentProperties.Company = "e-iceblue";
  document.BuiltinDocumentProperties.Manager = "Jakson";
  document.BuiltinDocumentProperties.Category = "Doc Demos";
  document.BuiltinDocumentProperties.Keywords = "Document, Property, Demo";
  document.BuiltinDocumentProperties.Comments = "This document is just a demo.";

  // Clean up resources
  document.Dispose();
}
```

---

# spire.doc javascript document operations
## load and save document to disk
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Initialize a new instance of WebClient class.
const response = await fetch("http://www.e-iceblue.com/images/test.docx");

// Download a Word document from URL.
let buffer = response.arrayBuffer();
let ms = wasmModule.Stream.CreateByBytes(buffer);
document.LoadFromStream({ stream: ms, fileFormat: wasmModule.FileFormat.Docx });

// Define the output file name
const outputFileName = "DownloadWordFromURL_out.docx";

// Save the document to the specified path
document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

// Clean up resources
document.Dispose();
```

---

# Document Stream Operations
## Load document from stream and save to stream in different format
```javascript
// Create a new document
let stream = wasmModule.Stream.CreateByFile(inputFileName);

// Load the entire document into memory.
let doc = wasmModule.Document.Create();
doc.LoadFromStream({ stream: stream, fileFormat: wasmModule.FileFormat.Auto })

// You can close the stream now, it is no longer needed because the document is in memory.
stream.Close();

// Convert the document to a different format and save to stream.
let newStream = wasmModule.Stream.CreateByFile(outputFileName);
doc.SaveToStream({ stream: newStream, fileFormat: wasmModule.FileFormat.Rtf });

// Clean up resources
doc.Dispose();
```

---

# Spire.Doc JavaScript Document Object Recursion
## Recursively traverse all document objects in a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

// Find all document objects
let builder = [];
for (let i = 0; i < document.Sections.Count; i++) {
  let section = document.Sections.get_Item(i);
  builder.push("section index " + i + " has following ChildObjects\n");

  for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
    let obj = section.Body.ChildObjects.get(j);
    builder.push("Index : " + j + ", ChildObject Type: " + obj.DocumentObjectType + "\n");
    if (obj instanceof wasmModule.Paragraph) {
      let paragraph = obj;
      builder.push("\tParagraph index " + section.Body.GetIndex(paragraph) + " has following ChildObjects\n");
      for (let k = 0; k < paragraph.ChildObjects.Count; k++) {
        let obj2 = paragraph.ChildObjects.get(k);
        builder.push("\tIndex : " + paragraph.GetIndex(obj2) + ", ChildObject Type: " + obj2.DocumentObjectType + "\n");
      }
    }
  }
  builder.push(" \n");
}
```

---

# Spire.Doc JavaScript View Modes
## Set Word document view modes
```javascript
// Create a new document
const document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

//Set Word view modes.
document.ViewSetup.DocumentViewType = wasmModule.DocumentViewType.WebLayout;
document.ViewSetup.ZoomPercent = 150;
document.ViewSetup.ZoomType = wasmModule.ZoomType.None;
```

---

# Spire.Doc JavaScript Update Document Property
## Update the last saved date of a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

// Function to convert local time to Greenwich time
function LocalTimeToGreenwishTime(localTime) {
  let localTimeZone = wasmModule.TimeZone.get_CurrentTimeZone();
  let timeSpan = localTimeZone.GetUtcOffset(localTime);
  let greenwishTime = localTime - timeSpan;
  return greenwishTime;
}

// Update the last saved date
document.BuiltinDocumentProperties.LastSaveDate = LocalTimeToGreenwishTime(wasmModule.DateTime.get_Now());

// Save the document to the specified path
document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

// Clean up resources
document.Dispose();
```

---

# Spire.Doc JavaScript Document Operation
## Add Section from Another Document
```javascript
// Create a new document
const TarDoc = wasmModule.Document.Create();
TarDoc.LoadFromFile(inputFileName1);

// Open a Word document as source document
let SouDoc = wasmModule.Document.Create();
SouDoc.LoadFromFile(inputFileName2);

// Get the second section from source document
let Ssection = SouDoc.Sections.get(1);

// Add the section in target document
TarDoc.Sections.Add(Ssection.Clone());
```

---

# Spire.doc JavaScript Document Cloning
## Clone a Word document using Spire.doc for JavaScript
```javascript
// Get the WASM module
wasmModule = window.wasmModule;

// Create and load a document
const document = wasmModule.Document.Create();
document.LoadFromFile("Template_Docx_1.docx");

// Clone the word document
let newDoc = document.Clone();

// Clean up resources
document.Dispose();
newDoc.Dispose();
```

---

# Spire.Doc JavaScript Document Operation
## Copy content from one Word document to another
```javascript
//Initialize a new object of Document class and load the source document.
let sourceDoc = wasmModule.Document.Create();
sourceDoc.LoadFromFile(inputFileName1);

//Initialize another object to load target document.
let destinationDoc = wasmModule.Document.Create();
destinationDoc.LoadFromFile(inputFileName2);

//Copy content from source file and insert them to the target file.
for (let i = 0; i < sourceDoc.Sections.Count; i++) {
  let sec = sourceDoc.Sections.get_Item(i);
  for (let j = 0; j < sec.Body.ChildObjects.Count; j++) {
    let obj = sec.Body.ChildObjects.get(j);
    destinationDoc.Sections.get(0).Body.ChildObjects.Add(obj.Clone());
  }
}
```

---

# Document Merge with Same Format
## Merge documents while preserving the format of the source document
```javascript
//Load the source document
let srcDoc = wasmModule.Document.Create();
srcDoc.LoadFromFile(inputFileName1);

//Load the destination document
let destDoc = wasmModule.Document.Create();
destDoc.LoadFromFile(inputFileName2);

//Keep same format of source document
srcDoc.KeepSameFormat = true;

//Copy the sections of source document to destination document
for (let i = 0; i < srcDoc.Sections.Count; i++) {
  let section = srcDoc.Sections.get_Item(i);
  destDoc.Sections.Add(section.Clone());
}
```

---

# Spire.Doc JavaScript Headers Footers
## Link headers and footers between Word document sections
```javascript
//Load the source file
let srcDoc = wasmModule.Document.Create()
srcDoc.LoadFromFile(inputFileName1);

//Load the destination file
let dstDoc = wasmModule.Document.Create();
dstDoc.LoadFromFile(inputFileName2);

//Link the headers and footers in the source file
srcDoc.Sections.get_Item(0).HeadersFooters.Header.LinkToPrevious = true;
srcDoc.Sections.get_Item(0).HeadersFooters.Footer.LinkToPrevious = true;

//Clone the sections of source to destination
for (let i = 0; i < srcDoc.Sections.Count; i++) {
  let section = srcDoc.Sections.get_Item(i);
  dstDoc.Sections.Add(section.Clone());
}
```

---

# Document Merge Operation
## Merge two Word documents by cloning sections
```javascript
// Load the first file
let document = wasmModule.Document.Create();
document.LoadFromFile({ fileName: inputFileName1, fileFormat: wasmModule.FileFormat.Doc });

// Load the second file
let documentMerge = wasmModule.Document.Create();
documentMerge.LoadFromFile({ fileName: inputFileName2, fileFormat: wasmModule.FileFormat.Docx });

// Merge documents
for (let i = 0; i < documentMerge.Sections.Count; i++) {
  let section = documentMerge.Sections.get_Item(i);
  document.Sections.Add(section.Clone());
}

// Save the document to the specified path
document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });
```

---

# Document Merge on Same Page
## Merge Word documents by cloning sections and content to the same page
```javascript
//Create a document
let document = wasmModule.Document.Create();

//Load the source document
document.LoadFromFile(inputFileName1);

//Clone a destination document
let destinationDocument = wasmModule.Document.Create();

//Load the destination document
destinationDocument.LoadFromFile(inputFileName2);

//Traverse sections
for (let i = 0; i < document.Sections.Count; i++) {
  let section = document.Sections.get_Item(i);
  //Traverse body ChildObjects
  for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
    let obj = section.Body.ChildObjects.get(j);
    //Clone to destination document at the same page
    destinationDocument.Sections.get(0).Body.ChildObjects.Add(obj.Clone());
  }
}
```

---

# Document Theme Preservation
## Preserve theme when copying sections from one Word document to another
```javascript
//Load the source document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Create a new Word document
let newWord = wasmModule.Document.Create();

//Clone default style, theme, compatibility from the source file to the destination document
doc.CloneDefaultStyleTo(newWord);
doc.CloneThemesTo(newWord);
doc.CloneCompatibilityTo(newWord);

//Add the cloned section to destination document
newWord.Sections.Add(doc.Sections.get_Item(0).Clone());
```

---

# Spire.Doc JavaScript Section Break
## Set section break as continuous in Word document
```javascript
// Open a Word document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

for (let i = 0; i < doc.Sections.Count; i++) {
  let section = doc.Sections.get_Item(i);
  // Set section break as continuous
  section.BreakCode = wasmModule.SectionBreakType.NoBreak;
}

// Save the document to the specified path
doc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

// Clean up resources
doc.Dispose();
```

---

# Spire.Doc JavaScript Document Operation
## Insert one Word document into another
```javascript
// Load the Word document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName1);

// Insert document from file
doc.InsertTextFromFile(inputFileName2, wasmModule.FileFormat.Auto);

// Save the document to the specified path
doc.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

// Clean up resources
doc.Dispose();
```

---

# Spire.Doc JavaScript Document Splitting
## Split a Word document into multiple documents by page breaks
```javascript
//Create Word document.
let original = wasmModule.Document.Create();
original.LoadFromFile(inputFileName);

//Create a new word document and add a section to it.
let newWord = wasmModule.Document.Create();
let section = newWord.AddSection();
original.CloneDefaultStyleTo(newWord);
original.CloneThemesTo(newWord);
original.CloneCompatibilityTo(newWord);

//Split the original word document into separate documents according to page break.
let index = 0;

//Traverse through all sections of original document.
for (let i = 0; i < original.Sections.Count; i++) {
  let sec = original.Sections.get(i);
  //Traverse through all body child objects of each section.
  for (let j = 0; j < sec.Body.ChildObjects.Count; j++) {
    let obj = sec.Body.ChildObjects.get(j);
    if (obj instanceof wasmModule.Paragraph) {
      let para = obj;
      sec.CloneSectionPropertiesTo(section);
      //Add paragraph object in original section into section of new document.
      section.Body.ChildObjects.Add(para.Clone());

      for (let k = 0; k < para.ChildObjects.Count; k++) {
        let parobj = para.ChildObjects.get(k);
        if (parobj instanceof wasmModule.Break && parobj.BreakType == wasmModule.BreakType.PageBreak) {
          //Get the index of page break in paragraph.
          let i = para.ChildObjects.IndexOf(parobj);

          //Remove the page break from its paragraph.
          section.Body.LastParagraph.ChildObjects.RemoveAt(i);

          //Create a new document and add a section.
          newWord = wasmModule.Document.Create();
          section = newWord.AddSection();
          original.CloneDefaultStyleTo(newWord);
          original.CloneThemesTo(newWord);
          original.CloneCompatibilityTo(newWord);
          sec.CloneSectionPropertiesTo(section);
          //Add paragraph object in original section into section of new document.
          section.Body.ChildObjects.Add(para.Clone());
          if (section.Paragraphs.get_Item(0).ChildObjects.Count == 0) {
            //Remove the first blank paragraph.
            section.Body.ChildObjects.RemoveAt(0);
          } else {
            //Remove the child objects before the page break.
            while (i >= 0) {
              section.Paragraphs.get_Item(0).ChildObjects.RemoveAt(i);
              i--;
            }
          }
        }
      }
    }
    if (obj instanceof wasmModule.Table) {
      //Add table object in original section into section of new document.
      section.Body.ChildObjects.Add(obj.Clone());
    }
  }
}
```

---

# spire.doc javascript document operation
## split document by section break
```javascript
//Split a Word document into multiple documents by section break.
for (let i = 0; i < document.Sections.Count; i++) {
  newWord = wasmModule.Document.Create();
  newWord.Sections.Add(document.Sections.get(i).Clone());
}
```

---

# spire.doc document splitting
## split document into multiple html pages
```javascript
function IsInNextDocument(element) {
  if (element instanceof wasmModule.Paragraph) {
    let p = element;
    if (p.StyleName == "Heading1") {
      return true;
    }
  }
  return false;
}

function SplitDocIntoMultipleHtml(input, outDirectory) {
  //Load file
  let document = wasmModule.Document.Create();
  document.LoadFromFile(input);

  //Create a new document
  let subDoc = wasmModule.Document.Create();
  subDoc.AddSection();
  let first = true;
  let index = 0;
  for (let i = 0; i < document.Sections.Count; i++) {
    let sec = document.Sections.get_Item(i);
    for (let j = 0; j < sec.Body.ChildObjects.Count; j++) {
      let element = sec.Body.ChildObjects.get(j);
      if (IsInNextDocument(element)) {
        if (!first) {
          //Embed css style and image data into html page
          subDoc.HtmlExportOptions.CssStyleSheetType = wasmModule.CssStyleSheetType.Internal;
          subDoc.HtmlExportOptions.ImageEmbedded = true;
          //Save to html file
          subDoc.SaveToFile({ fileName: outDirectory + `SplitDocIntoHtmlPages-${index}.docx`, fileFormat: wasmModule.FileFormat.Html });
          index++;
        }
        first = false;
      }
      if (subDoc == null) {
        subDoc = wasmModule.Document.Create();
        subDoc.AddSection();
      }
      subDoc.Sections.get(0).Body.ChildObjects.Add(element.Clone());
    }
  }
  if (subDoc != null) {
    //Embed css style and image data into html page
    subDoc.HtmlExportOptions.CssStyleSheetType = wasmModule.CssStyleSheetType.Internal;
    subDoc.HtmlExportOptions.ImageEmbedded = true;
    //Save to html file
    subDoc.SaveToFile({ fileName: outDirectory + `SplitDocIntoHtmlPages-${index}.docx`, fileFormat: wasmModule.FileFormat.Html });
    index++;
  }
  subDoc.Close();
  document.Close();
}
```

---

# Spire.Doc JavaScript Track Changes
## Accept or reject tracked changes in Word document
```javascript
//Get the first section and the paragraph we want to accept/reject the changes.
let sec = document.Sections.get(0);
let para = sec.Paragraphs.get_Item(0);

//Accept the changes or reject the changes.
para.Document.AcceptChanges();
//para.Document.RejectChanges();
```

---

# Spire.Doc JavaScript Track Changes
## Enable track changes in a Word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file 
document.LoadFromFile(inputFileName);

//Enable the track changes.
document.TrackChanges = true;

// Clean up resources
document.Dispose();
```

---

# Spire.Doc JavaScript Document Revisions
## Extract and analyze document revisions (track changes) from a Word document
```javascript
//Create a new document
let document = wasmModule.Document.Create();

//Load the file
document.LoadFromFile(inputFileName);

let insertRevision = []
insertRevision.push("Insert revisions:\n");
let index_insertRevision = 0;
let deleteRevision = [];
deleteRevision.push("Delete revisions:\n");
let index_deleteRevision = 0;
//Traverse sections
for (let i = 0; i < document.Sections.Count; i++) {
  let sec = document.Sections.get_Item(i);
  //Iterate through the element under body in the section
  for (let j = 0; j < sec.Body.ChildObjects.Count; j++) {
    let docItem = sec.Body.ChildObjects.get(j);
    if (docItem instanceof wasmModule.Paragraph) {
      let para = docItem;
      //Determine if the paragraph is an insertion revision
      if (para.IsInsertRevision) {
        index_insertRevision++;
        insertRevision.push("Index: " + index_insertRevision + "\n");
        //Get insertion revision
        let insRevison = para.InsertRevision;

        //Get insertion revision type
        let insType = insRevison.Type;
        insertRevision.push("Type: " + insType + "\n");
        //Get insertion revision author
        let insAuthor = insRevison.Author;
        insertRevision.push("Author: " + insAuthor + "\n");
      }
      //Determine if the paragraph is a delete revision
      else if (para.IsDeleteRevision) {
        index_deleteRevision++;
        deleteRevision.push("Index: " + index_deleteRevision + "\n");
        let delRevison = para.DeleteRevision;
        let delType = delRevison.Type;
        deleteRevision.push("Type: " + delType + "\n");
        let delAuthor = delRevison.Author;
        deleteRevision.push("Author: " + delAuthor + "\n");
      }
      //Iterate through the element in the paragraph
      for (let i = 0; i < para.ChildObjects.Count; i++) {
        let obj = para.ChildObjects.get(i);
        if (obj instanceof wasmModule.TextRange) {
          let textRange = obj;
          //Determine if the textrange is an insertion revision
          if (textRange.IsInsertRevision) {
            index_insertRevision++;
            insertRevision.push("Index: " + index_insertRevision + "\n");
            let insRevison = textRange.InsertRevision;
            let insType = insRevison.Type;
            insertRevision.push("Type: " + insType + "\n");
            let insAuthor = insRevison.Author;
            insertRevision.push("Author: " + insAuthor + "\n");
          } else if (textRange.IsDeleteRevision) {
            index_deleteRevision++;
            deleteRevision.push("Index: " + index_deleteRevision + "\n");
            //Determine if the textrange is a delete revision
            let delRevison = textRange.DeleteRevision;
            let delType = delRevison.Type;
            deleteRevision.push("Type: " + delType + "\n");
            let delAuthor = delRevison.Author;
            deleteRevision.push("Author: " + delAuthor + "\n");
          }
        }
      }
    }
  }
}
// Clean up resources
document.Dispose();
```

---

# Spire.Doc JavaScript Variables
## Add variables to a Word document
```javascript
// Create Word document
let document = wasmModule.Document.Create();

// Add a section
let section = document.AddSection();

// Add a paragraph
let paragraph = section.AddParagraph();

// Add a DocVariable field
paragraph.AppendField("A1", wasmModule.FieldType.FieldDocVariable);

// Add a document variable to the DocVariable field
document.Variables.Add("A1", "12");

// Update fields
document.IsUpdateFields = true;
```

---

# Word Document Variables Counter
## Count the number of variables in a Word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file from disk.
document.LoadFromFile(inputFileName);

//Get the number of variables in the document.
let number = document.Variables.Count;

let content = [];
content.push("The number of variables is: " + number.toString());
```

---

# Spire.Doc JavaScript Variables
## Get variables from a Word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file.
document.LoadFromFile(inputFileName);
let stringBuilder = [];
stringBuilder.push("This document has following variables:\n");
for (let i = 0; i < document.Variables.Count; i++) {
  let name = document.Variables.GetNameByIndex(i);
  let value = document.Variables.GetValueByIndex(i);
  stringBuilder.push("Name: " + name + ", " + "Value: " + value + "\n");
}
```

---

# spire.doc javascript variables
## remove variables from word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file from disk.
document.LoadFromFile(inputFileName);

//Remove the variable by name.
document.Variables.Remove("A1");
document.IsUpdateFields = true;
```

---

# spire.doc javascript variables
## retrieve variables from word document
```javascript
//Create Word document
let document = wasmModule.Document.Create();

//Load the file from disk
document.LoadFromFile(inputFileName);

//Retrieve name of the variable by index
let s1 = document.Variables.GetNameByIndex(0);

//Retrieve value of the variable by index
let s2 = document.Variables.GetValueByIndex(0);

//Retrieve the value of the variable by name
let s3 = document.Variables.get_Item("A1");

let content = [];
content.push("The name of the variable retrieved by index 0 is: " + s1 + "\n");
content.push("The vaule of the variable retrieved by index 0 is: " + s2 + "\n");
content.push("The vaule of the variable retrieved by name \"A1\" is: " + s3 + "\n");
```

---

# Spire.Doc JavaScript Gradient Background
## Set gradient background for Word document
```javascript
//Set the background type as Gradient.
document.Background.Type = wasmModule.BackgroundType.Gradient;
let Test = document.Background.Gradient;

//Set the first color and second color for Gradient.
Test.Color1 = wasmModule.Color.get_White();
Test.Color2 = wasmModule.Color.get_LightBlue();

//Set the Shading style and Variant for the gradient.
Test.ShadingVariant = wasmModule.GradientShadingVariant.ShadingDown;
Test.ShadingStyle = wasmModule.GradientShadingStyle.Horizontal;
```

---

# Spire.Doc JavaScript Background
## Set image background for Word document
```javascript
// Load a word document
let document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName1);

// Set the background type as picture
document.Background.Type = wasmModule.BackgroundType.Picture;

// Set the background picture
document.Background.SetPicture(inputFileName2);
```

---

# Spire.Doc JavaScript Page Setup
## Add gutter to Word document
```javascript
//Create Word document
let document = wasmModule.Document.Create();

//Load the file
document.LoadFromFile(inputFileName);

//Create a new section
let section = document.Sections.get(0);

//Set gutter
section.PageSetup.Gutter = 100;
```

---

# Word Document Line Numbering Setup
## Configure line numbering properties in Word document sections
```javascript
//Create Word document
let document = wasmModule.Document.Create();

//Load the file 
document.LoadFromFile(inputFileName);

//Set the start value of the line numbers
document.Sections.get(0).PageSetup.LineNumberingStartValue = 1;

//Set the interval between displayed numbers
document.Sections.get(0).PageSetup.LineNumberingStep = 6;

//Set the distance between line numbers and text
document.Sections.get(0).PageSetup.LineNumberingDistanceFromText = 40;

//Set the numbering mode of line numbers. There are four choices: None, Continuous, RestartPage and RestartSection
document.Sections.get(0).PageSetup.LineNumberingRestartMode = wasmModule.LineNumberingRestartMode.Continuous;
```

---

# spire.doc javascript page setup
## add page borders to word document
```javascript
// Get the first section
let section = document.Sections.get(0);

// Set the border type for the page setup of the section to Wave
section.PageSetup.Borders.BorderType = wasmModule.BorderStyle.Wave;

// Set the color of the borders to Green
section.PageSetup.Borders.Color = wasmModule.Color.get_Green();

// Set the left spacing for the borders of the page setup
section.PageSetup.Borders.Left.Space = 20;

// Set the right spacing for the borders of the page setup
section.PageSetup.Borders.Right.Space = 20;
```

---

# Spire.Doc JavaScript Page Setup
## Add page numbers in different sections of a Word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Repeat step2 and Step3 for the rest sections, so change the code with for loop.
for (let i = 0; i < 3; i++) {
  let footer = document.Sections.get(i).HeadersFooters.Footer;
  let footerParagraph = footer.AddParagraph();
  footerParagraph.AppendField("page number", wasmModule.FieldType.FieldPage);
  footerParagraph.AppendText(" of ");
  footerParagraph.AppendField("number of pages", wasmModule.FieldType.FieldSectionPages);
  footerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

  if (i == 2)
    break;
  else {
    document.Sections.get(i + 1).PageSetup.RestartPageNumbering = true;
    document.Sections.get(i + 1).PageSetup.PageStartingNumber = 1;
  }
}
```

---

# Word Document Page Setup
## Configure different page setup options for Word document sections
```javascript
//Get the second section
let SectionTwo = document.Sections.get(1);

//Set the orientation
SectionTwo.PageSetup.Orientation = wasmModule.PageOrientation.Landscape;

//Set page size
//SectionTwo.PageSetup.PageSize = new SizeF(800, 800);
```

---

# Spire.Doc JavaScript Section Break
## Insert section break in Word document

```javascript
// Create word document
let document = wasmModule.Document.Create();

// Add section
let section = document.AddSection();

// Page setup
SetPage(section);

// Add cover
InsertCover(section);

// Insert a break code
section = document.AddSection();
section.AddParagraph().InsertSectionBreak({ breakType: wasmModule.SectionBreakType.NewPage });

// Add content
InsertContent(section);
```

---

# Insert Page Break in Word Document
## Find specific text and insert page break after it
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file from disk.
document.LoadFromFile(inputFileName);

//Find the specified word "technology" where we want to insert the page break.
let selections = document.FindAllString("technology", true, true);

//Traverse each word "technology".
for (let ts of selections) {
  let range = ts.GetAsOneRange();
  let paragraph = range.OwnerParagraph;
  let index = paragraph.ChildObjects.IndexOf(range);

  //Create a new instance of page break and insert a page break after the word "technology".
  let pageBreak = wasmModule.Break.Create(document, wasmModule.BreakType.PageBreak);
  paragraph.ChildObjects.Insert(index + 1, pageBreak);
}
```

---

# Spire.Doc JavaScript Page Break
## Insert page break in Word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file
document.LoadFromFile(inputFileName);

//Insert page break.
document.Sections.get(0).Paragraphs.get_Item(3).AppendBreak(wasmModule.BreakType.PageBreak);
```

---

# Word Document Section Break Insertion
## Insert section break in Word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file from disk.
document.LoadFromFile(inputFileName);

//Insert section break. There are five section break options including EvenPage, NewColumn, NewPage, NoBreak, OddPage.
document.Sections.get(0).Paragraphs.get_Item(1).InsertSectionBreak({ breakType: wasmModule.SectionBreakType.NoBreak });

// Define the output file name
const outputFileName = "InsertSectionBreak_out.docx";

// Save the document to the specified path
document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

// Clean up resources
document.Dispose();
```

---

# Word Document Page Setup
## Configure page properties including size, margins, headers, and footers in a Word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();
let section = document.AddSection();

//The unit of all measures below is point, 1point = 0.3528 mm.
section.PageSetup.PageSize = wasmModule.PageSize.A4();
section.PageSetup.Margins.Top = 72;
section.PageSetup.Margins.Bottom = 72;
section.PageSetup.Margins.Left = 89.85;
section.PageSetup.Margins.Right = 89.85;

function InsertHeaderAndFooter(section, headerPic, footerPic) {
  let header = section.HeadersFooters.Header;
  let footer = section.HeadersFooters.Footer;

  //Insert picture and text to header.
  let headerParagraph = header.AddParagraph();
  let headerPicture = headerParagraph.AppendPicture({ imgFile: headerPic });

  //Header text.
  let text = headerParagraph.AppendText("Demo of Spire.Doc");
  text.CharacterFormat.FontName = "Arial";
  text.CharacterFormat.FontSize = 10;
  text.CharacterFormat.Italic = true;
  headerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

  //Border.
  headerParagraph.Format.Borders.Bottom.BorderType = wasmModule.BorderStyle.Single;
  headerParagraph.Format.Borders.Bottom.Space = 0.05;

  //Header picture layout - text wrapping.
  headerPicture.TextWrappingStyle = wasmModule.TextWrappingStyle.Behind;

  //Header picture layout - position.
  headerPicture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
  headerPicture.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
  headerPicture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
  headerPicture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Top;

  //Insert picture to footer.
  let footerParagraph = footer.AddParagraph();
  let footerPicture = footerParagraph.AppendPicture({ imgFile: footerPic });

  //Footer picture layout.
  footerPicture.TextWrappingStyle = wasmModule.TextWrappingStyle.Behind;
  footerPicture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
  footerPicture.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
  footerPicture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
  footerPicture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Bottom;

  //Insert page number.
  footerParagraph.AppendField("page number", wasmModule.FieldType.FieldPage);
  footerParagraph.AppendText(" of ");
  footerParagraph.AppendField("number of pages", wasmModule.FieldType.FieldNumPages);
  footerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

  //Border.
  footerParagraph.Format.Borders.Top.BorderType = wasmModule.BorderStyle.Single;
  footerParagraph.Format.Borders.Top.Space = 0.05;
}
```

---

# Spire.Doc JavaScript Page Setup
## Remove page breaks from Word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file from disk.
document.LoadFromFile(inputFileName);

//Traverse every paragraph of the first section of the document.
for (let j = 0; j < document.Sections.get(0).Paragraphs.Count; j++) {
  let p = document.Sections.get(0).Paragraphs.get_Item(j);

  //Traverse every child object of a paragraph.
  for (let i = 0; i < p.ChildObjects.Count; i++) {
    let obj = p.ChildObjects.get(i);

    //Find the page break object.
    if (obj.DocumentObjectType == wasmModule.DocumentObjectType.Break) {
      let b = obj;

      //Remove the page break object from paragraph.
      p.ChildObjects.Remove(b);
    }
  }
}

// Save the document to the specified path
document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });
```

---

# Spire.Doc JavaScript Page Setup
## Reset page number for each section to start at 1
```javascript
// Create three Word documents and load three different Word documents 
let document1 = wasmModule.Document.Create();
document1.LoadFromFile(inputFileName1);

let document2 = wasmModule.Document.Create();
document2.LoadFromFile(inputFileName2);

let document3 = wasmModule.Document.Create();
document3.LoadFromFile(inputFileName3);

// Use section method to combine all documents into one word document
for (let i = 0; i < document2.Sections.Count; i++) {
  let sec = document2.Sections.get(i);
  document1.Sections.Add(sec.Clone());
}
for (let i = 0; i < document3.Sections.Count; i++) {
  let sec = document3.Sections.get(i);
  document1.Sections.Add(sec.Clone());
}

// Traverse every section of document1
for (let i = 0; i < document1.Sections.Count; i++) {
  let sec = document1.Sections.get_Item(i);
  // Traverse every object of the footer
  for (let j = 0; j < sec.HeadersFooters.Footer.ChildObjects.Count; j++) {
    let obj = sec.HeadersFooters.Footer.ChildObjects.get(j);
    if (obj.DocumentObjectType == wasmModule.DocumentObjectType.StructureDocumentTag) {
      let para = obj.ChildObjects.get(0);
      for (let k = 0; k < para.ChildObjects.Count; k++) {
        let item = para.ChildObjects.get(k);
        if (item.DocumentObjectType == wasmModule.DocumentObjectType.Field)
          // Find the item and its field type is FieldNumPages
          if (item.Type == wasmModule.FieldType.FieldNumPages) {
            // Change field type to FieldSectionPages
            item.Type = FieldType.FieldSectionPages;
          }
      }
    }
  }
}

// Restart page number of section and set the starting page number to 1
document1.Sections.get(1).PageSetup.RestartPageNumbering = true;
document1.Sections.get(1).PageSetup.PageStartingNumber = 1;

document1.Sections.get(2).PageSetup.RestartPageNumbering = true;
document1.Sections.get(2).PageSetup.PageStartingNumber = 1;
```

---

# Document to Byte Conversion
## Convert Word document to byte array and back to document
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Create a new memory stream.
let outStream = wasmModule.Stream.Create();
// Save the document to stream.
doc.SaveToStream({stream: outStream, fileFormat: wasmModule.FileFormat.Docx});

// Convert the document to bytes.
let docBytes = outStream.ToArray();

// Now reverse the steps to load the bytes back into a document object.
let inStream = wasmModule.Stream.CreateByBytes(docBytes);

// Load the stream into a new document object.
let newDoc = wasmModule.Document.Create();
newDoc.LoadFromStream({stream: inStream, fileFormat: wasmModule.FileFormat.Auto});
```

---

# HTML to Image Conversion
## Convert HTML document to image format
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load the file from disk
doc.LoadFromFile({fileName: inputFileName, fileFormat: wasmModule.FileFormat.Html, validationType: wasmModule.XHTMLValidationType.None});

// Save to image. You can convert HTML to BMP, JPEG, PNG, GIF, Tiff, etc.
let image = doc.SaveImageToStreams({pageIndex: 0, imagetype: wasmModule.ImageType.Bitmap});

// Define the output file name
const outputFileName = "HtmlToImage-result.png";
image.Save(outputFileName);

// Clean up resources
doc.Dispose();
```

---

# spire.doc javascript html to pdf conversion
## convert html file to pdf format
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load the HTML file
doc.LoadFromFile({fileName:inputFileName, fileFormat : wasmModule.FileFormat.Html, validationType:wasmModule.XHTMLValidationType.None});

// Save as PDF
doc.SaveToFile({fileName:outputFileName, fileFormat:wasmModule.FileFormat.PDF});

// Clean up resources
doc.Dispose();
```

---

# HTML to XML Conversion
## Convert HTML file to XML format using Spire.Doc for JavaScript
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile({fileName:inputFileName, fileFormat : wasmModule.FileFormat.Html, validationType:wasmModule.XHTMLValidationType.None});

// Define the output file name
const outputFileName = "HtmlToXml-result.xml";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Xml});
```

---

# HTML to XPS Conversion
## Convert HTML document to XPS format using Spire.Doc for JavaScript
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile({
  fileName: inputFileName, 
  fileFormat: wasmModule.FileFormat.Html, 
  validationType: wasmModule.XHTMLValidationType.None
});

// Define the output file name
const outputFileName = "HtmlToXps-result.xps";

// Save the document to the specified path
doc.SaveToFile({
  fileName: outputFileName, 
  fileFormat: wasmModule.FileFormat.XPS
});

// Clean up resources
doc.Dispose();
```

---

# spire.doc javascript image to pdf conversion
## convert image to PDF document
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Create a new section
let section = doc.AddSection();
// Create a new paragraph
let paragraph = section.AddParagraph();
// Add a picture for paragraph
paragraph.AppendPicture({imgFile: inputFileName});

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.PDF});

// Clean up resources
doc.Dispose();
```

---

# Spire.Doc JavaScript ODT to Word Conversion
## Convert ODT files to DOCX format
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "OdtToWord-result.docx";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx});

// Clean up resources
doc.Dispose();
```

---

# RTF to HTML Conversion
## Convert RTF document to HTML format
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Html});

// Clean up resources
doc.Dispose();
```

---

# RTF to PDF Conversion
## Converting RTF documents to PDF format using Spire.Doc for JavaScript
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "RtfToPdf-result.pdf";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.PDF});

// Clean up resources
doc.Dispose();
```

---

# Spire.Doc JavaScript Conversion
## Convert Word document to image
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document
doc.LoadFromFile(inputFileName);

// Convert document to image
let img = doc.SaveImageToStreams({pageIndex: 0, imagetype: wasmModule.ImageType.Bitmap});

// Clean up resources
doc.Dispose();
```

---

# Word to ODT Conversion
## Convert Word document to ODT format using Spire.Doc for JavaScript
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToOdt-result.odt";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Odt});

// Clean up resources
doc.Dispose();
```

---

# Word to PCL Conversion
## Convert Word document to PCL (Printer Command Language) format
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToPCL-result.pcl";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.PCL});

// Clean up resources
doc.Dispose();
```

---

# Word to PostScript Conversion
## Convert Word document to PostScript format
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToPostScript-result.ps";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.PostScript});

// Clean up resources
doc.Dispose();
```

---

# Word to RTF Conversion
## Convert Word document to RTF format using Spire.Doc for JavaScript
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToRtf-result.rtf";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Rtf});

// Clean up resources
doc.Dispose();
```

---

# spire.doc javascript conversion
## convert Word document to SVG format
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile("input.docx");

// Save the document to SVG format
doc.SaveToFile({fileName: "output.svg", fileFormat: wasmModule.FileFormat.SVG});

// Clean up resources
doc.Dispose();
```

---

# spire.doc javascript conversion
## convert Word document to XML format
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToXML-result.xml";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Xml});

// Clean up resources
doc.Dispose();
```

---

# Word to XPS Conversion
## Convert Word document to XPS format using JavaScript
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToXPS-result.xps";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.XPS});

// Clean up resources
doc.Dispose();
```

---

# TXT to Word Conversion
## Convert a text file to a Word document using JavaScript
```javascript
// Define the input file name
let inputFileName = "TxtToWord.txt";

// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "TxtToWord-result.docx";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

// Clean up resources
doc.Dispose();
```

---

# Word to PDF/A Conversion
## Convert Word document to PDF/A format using JavaScript
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Set the Conformance-level of the Pdf file to PDF_A1B.
let toPdf = wasmModule.ToPdfParameterList.Create();
toPdf.PdfConformanceLevel = wasmModule.PdfConformanceLevel.Pdf_A1B;

// Save the file.
doc.SaveToFile({fileName: outputFileName, fileFormat: toPdf});

// Clean up resources
doc.Dispose();
```

---

# spire.doc javascript conversion
## convert Word document to txt file
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "WordToTxt-result.txt";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Txt});

// Clean up resources
doc.Dispose();
```

---

# Word to Word XML Conversion
## Convert Word document to Word XML format using JavaScript
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "WordToWordXML-result.xml";

// For word 2003:
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.WordML});
// //For word 2007:
// doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.WordXml});

// Clean up resources
doc.Dispose();
```

---

# XML to PDF Conversion
## Convert XML document to PDF format
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "XMLToPDF-result.pdf";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.PDF});

// Clean up resources
doc.Dispose();
```

---

# XML to Word Conversion
## Convert XML file to Word document using Spire.Doc
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "XMLToWord-result.docx";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

// Clean up resources
doc.Dispose();
```

---

# HTML to Word Conversion
## Convert HTML file to Word document using JavaScript
```javascript
// Create a new document
const document = wasmModule.Document.Create();

document.LoadFromFile({fileName: inputFileName,fileFormat: wasmModule.FileFormat.Html,validationType: wasmModule.XHTMLValidationType.None});

// Define the output file name
const outputFileName = "HtmlFileToWord-result.docx";

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx});

// Clean up resources
document.Dispose();
```

---

# Spire.Doc JavaScript HTML to Word Conversion
## Convert HTML string to Word document
```javascript
// HTML string to convert
let HTML = "<html><head><style type=\"text/css\">li,p{font-family:'Lucida Sans Unicode';font-size:14pt;}</style></head><body><font size=\"16pt\" color=\"blue\"><h2 align=\"center\">Spire.Doc</h2></font><p><b>Edition:</b></p>";
HTML+="<ul type=\"disc\"><li><span style='color:green;'>Free Edition</span></li><li>Trial version</li><li><span style='color:red'>A month free for trial version</span></li></ul></ul><p><b>Platform:</b></p><ul type=\"square\">";
HTML+="<li>.NET</li><li>WPF</li><li>Silverlight</li></ul><table border=\"1\" width=\"90%\"><tr><th>Main Functions of Spire.Doc</th></tr><tr><td>Convert File Documents with High Quality</td></tr> <tr><td>Richest Word Document Features Support</td></tr>";
HTML+="<tr><td>Simple & Easy to Process Pre-Existing Word Documents</td></tr><tr><td>Other Technical Features</td></tr></table></body></html>";

// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Create a new paragraph
let paragraph = section.AddParagraph();

// Append html string
paragraph.AppendHTML(HTML);

// Save the document
document.SaveToFile({fileName: "HtmlStringToWord-result.docx", fileFormat: wasmModule.FileFormat.Docx2013});

// Clean up resources
document.Dispose();
```

---

# Word to EPUB Conversion with Cover Image
## Convert Word document to EPUB format while adding a cover image
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName1);

let picture = wasmModule.DocPicture.Create(document);
picture.LoadImage(inputFileName2);

// Define the output file name
const outputFileName = "AddCoverImage-result.epub";

// Save the document to the specified path
document.SaveToEpub(outputFileName, picture);

// Clean up resources
document.Dispose();
```

---

# Word to EPUB Conversion
## Convert Word document to EPUB format
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToEpub-result.epub";
document.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.EPub});

// Clean up resources
document.Dispose();
```

---

# Word to HTML Conversion
## Convert Word document to HTML format
```javascript
//Create word document
let document = wasmModule.Document.Create();

document.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToHtml-result.html";

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Html});

// Clean up resources
document.Dispose();
```

---

# Word to HTML Conversion with Export Options
## Set various options when exporting Word document to HTML format
```javascript
// Open a Word document
let document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

// Set whether the css styles are embeded or not
document.HtmlExportOptions.CssStyleSheetFileName = outputDirectoryName + "sample.css";
document.HtmlExportOptions.CssStyleSheetType = wasmModule.CssStyleSheetType.External;

// Set whether the images are embeded or not
document.HtmlExportOptions.ImageEmbedded = false;
document.HtmlExportOptions.ImagesPath = outputDirectoryName + "Demo/";

// Set the option whether to export form fields as plain text or not
document.HtmlExportOptions.IsTextInputFormFieldAsText = true;

// Save the document to HTML format
document.SaveToFile({
  fileName: outputFileName,
  fileFormat: wasmModule.FileFormat.Html
});
```

---

# Word to PDF Conversion with Hyperlink Control
## Disable or preserve hyperlinks when converting Word documents to PDF
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load the file from disk
document.LoadFromFile(inputFileName);

// Create an instance of ToPdfParameterList
let pdf = wasmModule.ToPdfParameterList.Create();

// Set DisableLink to true to remove the hyperlink effect for the result PDF page
// Set DisableLink to false to preserve the hyperlink effect for the result PDF page
pdf.DisableLink = true;

// Save to file
document.SaveToFile({fileName: outputFileName, paramList: pdf});

// Clean up resources
document.Dispose();
```

---

# Word to PDF Conversion with Font Embedding
## Demonstrates how to convert Word to PDF and embed all fonts into PDF
```javascript
// Create a new document
const document = wasmModule.Document.Create();

document.LoadFromFile(inputFileName);

// Embeds full fonts by default when IsEmbeddedAllFonts is set to true.
let ppl = wasmModule.ToPdfParameterList.Create();
ppl.IsEmbeddedAllFonts = true;

// Define the output file name
const outputFileName = "EmbededAllFontsInPDF-result.pdf";

// Save doc file to pdf.
document.SaveToFile({fileName: outputFileName, paramList: ppl});

// Clean up resources
document.Dispose();
```

---

# Word to PDF Conversion with Embedded Fonts
## Convert Word document to PDF while embedding non-installed fonts
```javascript
// Create a new document
const document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

// Embed the non-installed fonts.
let parms = wasmModule.ToPdfParameterList.Create();
let fonts = [wasmModule.PrivateFontPath.Create("PT Serif Caption", "PT_Serif-Caption-Web-Regular.ttf")];
parms.PrivateFontPaths = fonts;

// Define the output file name
const outputFileName = "EmbedNoninstalledFonts-result.pdf";
// Save doc file to pdf.
document.SaveToFile({fileName: outputFileName, paramList: parms});

// Clean up resources
document.Dispose();
```

---

# spire.doc javascript conversion
## keep hidden text when converting Word to PDF
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load the file from disk
document.LoadFromFile(inputFileName);

// When convert to PDF file, set the property IsHidden as true
let pdf = wasmModule.ToPdfParameterList.Create();
pdf.IsHidden = true;

// Define the output file name
const outputFileName = "KeepHiddenText-result.pdf";

// Save to file
document.SaveToFile({fileName: outputFileName, paramList: pdf});

// Clean up resources
document.Dispose();
```

---

# Spire.Doc JavaScript Conversion
## Set image quality during Doc to PDF conversion
```javascript
// Create a new document
const document = wasmModule.Document.Create();

//Load the file from disk.
document.LoadFromFile("SetImageQuality.doc");

//Set the output image quality to be 40% of the original image. The default set of the output image quality is 80% of the original.
document.JPEGQuality = 40;

// Save the document to the specified path
document.SaveToFile({fileName: "SetImageQuality-result.pdf", fileFormat: wasmModule.FileFormat.PDF});

// Clean up resources
document.Dispose();
```

---

# Word to PDF Conversion with Embedded Fonts
## Convert Word document to PDF while specifying which fonts to embed
```javascript
// Create a new document
const document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

// Specify embedded font
let parms = wasmModule.ToPdfParameterList.Create();
let part = [];
part.push("PT Serif Caption");
parms.EmbeddedFontNameList = part;

// Define the output file name
const outputFileName = "SpecifyEmbeddedFont-result.pdf";

document.SaveToFile({fileName: outputFileName, paramList: parms});

// Clean up resources
document.Dispose();
```

---

# Word to PDF Conversion
## Convert Word document to PDF format using JavaScript
```javascript
// Create a new document
const document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "ToPDF-result.pdf";

// Save the document to a PDF file
document.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.PDF});

// Clean up resources
document.Dispose();
```

---

# Word to PDF Conversion with Bookmarks
## Convert Word document to PDF and create bookmarks
```javascript
// Create a new document
const document = wasmModule.Document.Create();
//Load the document from disk
document.LoadFromFile(inputFileName);

let parames = wasmModule.ToPdfParameterList.Create();
//Set CreateWordBookmarks to true
parames.CreateWordBookmarks = true;
parames.CreateWordBookmarksUsingHeadings = false;

// Define the output file name
const outputFileName = "ToPDFAndCreateBookmarks-result.pdf";

document.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.PDF});

// Clean up resources
document.Dispose();
```

---

# Word to PDF Conversion with Password
## Convert Word document to PDF with password protection
```javascript
// Create a new document
const document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

// Create a parameter
let toPdf = wasmModule.ToPdfParameterList.Create();

// Set the password
let password = "E-iceblue";
toPdf.PdfSecurity.Encrypt("password", password, wasmModule.PdfPermissionsFlags.Default, wasmModule.PdfEncryptionKeySize.Key128Bit);

// Define the output file name
const outputFileName = "ToPdfWithPassword-result.pdf";

// Save doc file
document.SaveToFile({fileName: outputFileName, paramList: toPdf});

// Clean up resources
document.Dispose();
```

---

# Change Font Color in Word Document
## Modify text color in Word document paragraphs
```javascript
// Get the first section and first paragraph
let section = document.Sections.get_Item(0);
let p1 = section.Paragraphs.get_Item(0);

// Iterate through the childObjects of the paragraph 1
for (let i = 0; i < p1.ChildObjects.Count; i++) {
    let childObj = p1.ChildObjects.get(i);
    if (childObj instanceof wasmModule.TextRange) {
        // Change text color
        let tr = childObj;
        tr.CharacterFormat.TextColor = wasmModule.Color.get_RosyBrown();
    }
}

// Get the second paragraph
let p2 = section.Paragraphs.get_Item(1);

// Iterate through the childObjects of the paragraph 2
for (let i = 0; i < p2.ChildObjects.Count; i++) {
    let childObj = p2.ChildObjects.get(i);
    if (childObj instanceof wasmModule.TextRange) {
        // Change text color
        let tr = childObj;
        tr.CharacterFormat.TextColor = wasmModule.Color.get_DarkGreen();
    }
}
```

---

# Embed Private Font in Word Document
## Demonstrates how to embed a private font into a Word document using Spire.Doc for JavaScript
```javascript
// Create a new document
const document = wasmModule.Document.Create();

//Get the first section and add a paragraph
let section = document.Sections.get_Item(0);
let p = section.AddParagraph();

//Append text to the paragraph, then set the font name and font size
let range = p.AppendText("Your Office Development Master");
range.CharacterFormat.FontName = "PT Serif Caption";
range.CharacterFormat.FontSize = 20;

//Allow embedding font in document
document.EmbedFontsInFile = true;

//Embed private font from font file into the document
document.AddPrivateFont(wasmModule.PrivateFontPath.Create("PT Serif Caption",  "PT Serif Caption.ttf"))

// Save the document
document.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013});
```

---

# Get List of Using Fonts
## Extract font information from a Word document
```javascript
let font_obj = new Map();

//Load word document
let document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

for (let i = 0; i < document.Sections.Count; i++) {
    let section = document.Sections.get_Item(i);
    for (let j = 0; j < section.Body.Paragraphs.Count; j++) {
        let paragraph = section.Body.Paragraphs.get_Item(j);
        for (let k = 0; k < paragraph.ChildObjects.Count; k++) {
            let obj = paragraph.ChildObjects.get(k);
            if (obj instanceof wasmModule.TextRange) {
                let range = obj;
                let font = {size:  range.CharacterFormat.FontSize, name: range.CharacterFormat.FontName};
                if (!font_obj.has(font)) {
                    font_obj.set(font, range);
                }
            }
        }
    }
}
for (let [Key,Value] of font_obj) {
  let font = Key;
  let range = Value;
  let s = "Font Name: " + font.name + ",Size: " + font.size  + ",Color: " + range.CharacterFormat.TextColor.Name;
  stringBuilder.push(s + "\n");
}
```

---

# Spire.Doc JavaScript Font Setting
## Set font properties in Word document
```javascript
//Get the first section
let s = document.Sections.get(0);

//Get the second paragraph
let p = s.Paragraphs.get_Item(1);

//Create a characterFormat object
let format = wasmModule.CharacterFormat.Create(document);
//Set font

format.FontName = "Arial";
format.FontSize = 16;
//Loop through the childObjects of paragraph
for (let i = 0; i < p.ChildObjects.Count; i++) {
    let childObj = p.ChildObjects.get(i);
    if (childObj instanceof wasmModule.TextRange) {
        //Apply character format
        let tr = childObj;
        tr.ApplyCharacterFormat(format);
    }
}
```

---

# ASCII Characters Bullet Style
## Create bullet styles using ASCII characters in a Word document
```javascript
//Create a new document
let document = wasmModule.Document.Create();
let section = document.AddSection();

//Create four list styles based on different ASCII characters
let listStyle1 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
listStyle1.Name = "liststyle";
listStyle1.Levels.get_Item(0).BulletCharacter = String.fromCharCode(0x006e);
listStyle1.Levels.get_Item(0).CharacterFormat.FontName = "Wingdings";
document.ListStyles.Add(listStyle1);
let listStyle2 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
listStyle2.Name = "liststyle2";
listStyle2.Levels.get_Item(0).BulletCharacter =String.fromCharCode(0x0075);
listStyle2.Levels.get_Item(0).CharacterFormat.FontName = "Wingdings";
document.ListStyles.Add(listStyle2);
let listStyle3 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
listStyle3.Name = "liststyle3";
listStyle3.Levels.get_Item(0).BulletCharacter = String.fromCharCode(0x00b2);
listStyle3.Levels.get_Item(0).CharacterFormat.FontName = "Wingdings";
document.ListStyles.Add(listStyle3);
let listStyle4 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
listStyle4.Name = "liststyle4";
listStyle4.Levels.get_Item(0).BulletCharacter = String.fromCharCode(0x00d8);
listStyle4.Levels.get_Item(0).CharacterFormat.FontName = "Wingdings";
document.ListStyles.Add(listStyle4);

//Add four paragraphs and apply list style separately
let p1 = section.Body.AddParagraph();
p1.AppendText("Spire.Doc for JavaScript");
p1.ListFormat.ApplyStyle(listStyle1.Name);
let p2 = section.Body.AddParagraph();
p2.AppendText("Spire.Doc for JavaScript");
p2.ListFormat.ApplyStyle(listStyle2.Name);
let p3 = section.Body.AddParagraph();
p3.AppendText("Spire.Doc for JavaScript");
p3.ListFormat.ApplyStyle(listStyle3.Name);
let p4 = section.Body.AddParagraph();
p4.AppendText("Spire.Doc for JavaScript");
p4.ListFormat.ApplyStyle(listStyle4.Name);
```

---

# Spire.Doc JavaScript Character Formatting
## Demonstrates how to apply various character formatting options to text in a Word document
```javascript
// Create a document
let document = wasmModule.Document.Create();
let sec = document.AddSection();
let titleParagraph = sec.AddParagraph();
titleParagraph.AppendText("Font Styles and Effects ");
titleParagraph.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Title});

let paragraph = sec.AddParagraph();
let tr = paragraph.AppendText("Strikethough Text");
tr.CharacterFormat.IsStrikeout = true;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Shadow Text");
tr.CharacterFormat.IsShadow = true;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Small caps Text");
tr.CharacterFormat.IsSmallCaps = true;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Double Strikethough Text");
tr.CharacterFormat.DoubleStrike = true;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Outline Text");
tr.CharacterFormat.IsOutLine = true;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("AllCaps Text");
tr.CharacterFormat.AllCaps = true;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Text");
tr = paragraph.AppendText("SubScript");
tr.CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SubScript;

tr = paragraph.AppendText("And");
tr = paragraph.AppendText("SuperScript");
tr.CharacterFormat.SubSuperScript = wasmModule.SubSuperScript.SuperScript;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Emboss Text");
tr.CharacterFormat.Emboss = true;
tr.CharacterFormat.TextColor = wasmModule.Color.get_White();

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Hidden:");
tr = paragraph.AppendText("Hidden Text");
tr.CharacterFormat.Hidden = true;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Engrave Text");
tr.CharacterFormat.Engrave = true;
tr.CharacterFormat.TextColor = wasmModule.Color.get_White();

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("WesternFonts中文字体");
tr.CharacterFormat.FontNameAscii = "Calibri";
tr.CharacterFormat.FontNameNonFarEast = "Calibri";
tr.CharacterFormat.FontNameFarEast = "Simsun";

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Font Size");
tr.CharacterFormat.FontSize = 20;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Font Color");
tr.CharacterFormat.TextColor = wasmModule.Color.get_Red();

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Bold Italic Text");
tr.CharacterFormat.Bold = true;
tr.CharacterFormat.Italic = true;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Underline Style");
tr.CharacterFormat.UnderlineStyle = wasmModule.UnderlineStyle.Single;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Highlight Text");
tr.CharacterFormat.HighlightColor = wasmModule.Color.get_Yellow();

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Text has shading");
tr.CharacterFormat.TextBackgroundColor = wasmModule.Color.get_Green();

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Border Around Text");
tr.CharacterFormat.Border.BorderType = wasmModule.BorderStyle.Single;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Text Scale");
tr.CharacterFormat.TextScale = 150;

paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
tr = paragraph.AppendText("Character Spacing is 2 point");
tr.CharacterFormat.CharacterSpacing = 2;
```

---

# Copy Document Styles
## Copy styles from one Word document to another
```javascript
//Load source document from disk
let srcDoc = wasmModule.Document.Create();
srcDoc.LoadFromFile(inputFileName_1);

//Load destination document from disk
let destDoc = wasmModule.Document.Create();
destDoc.LoadFromFile(inputFileName_2);

//Get the style collections of source document
let styles = srcDoc.Styles;

//Add the style to destination document
for (let i = 0; i < styles.Count; i++) {
    let style = styles.get_Item(i);
    destDoc.Styles.Add(style);
}
```

---

# spire.doc javascript character spacing
## get character spacing from word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Get the first section of document
let section = document.Sections.get(0);

// Get the first paragraph
let paragraph = section.Paragraphs.get_Item(0);

// Define two variables
let fontName = "";
let fontSpacing = 0;

// Traverse the ChildObjects
for (let i = 0; i < paragraph.ChildObjects.Count; i++) {
    let docObj = paragraph.ChildObjects.get(i);
    // If it is TextRange
    if (docObj instanceof wasmModule.TextRange) {
        let textRange = docObj;
        fontName = textRange.CharacterFormat.FontName;

        // Get the character spacing
        fontSpacing = textRange.CharacterFormat.CharacterSpacing;
    }
}
```

---

# Get Text by Style Name
## Extract text from Word document paragraphs with specific style name
```javascript
//Load document from disk
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Create string builder
let builder = []

//Loop through sections
for (let i = 0; i < doc.Sections.Count; i++) {
    let section = doc.Sections.get(i);
    //Loop through paragraphs
    for (let j = 0; j < section.Paragraphs.Count; j++) {
        let para = section.Paragraphs.get_Item(j);
        //Find the paragraph whose style name is "Heading1"
        if (para.StyleName == "Heading1") {
            //Write the text of paragraph
            builder.push(para.Text + "\n");
        }
    }
}

//Write the contents in a TXT file
wasmModule.FS.writeFile(outputFileName, builder.join("\n"));
doc.Close();
```

---

# spire.doc javascript list styles
## create and apply list styles to paragraphs
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Add a section
let sec = document.AddSection();

// Create numbered list style
let numberList = wasmModule.ListStyle.Create(document, wasmModule.ListType.Numbered);
numberList.Name = "numberList";
numberList.Levels.get_Item(1).NumberPrefix = "%1.";
numberList.Levels.get_Item(1).PatternType = wasmModule.ListPatternType.Arabic;
numberList.Levels.get_Item(2).NumberPrefix = "%1.%2.";
numberList.Levels.get_Item(2).PatternType = wasmModule.ListPatternType.Arabic;

// Create bulleted list style
let bulletList = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
bulletList.Name = "bulletList";

// Add the list styles into document
document.ListStyles.Add(numberList);
document.ListStyles.Add(bulletList);

// Add paragraph and apply the numbered list style
let paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 1");
paragraph.ListFormat.ApplyStyle(numberList.Name);

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2");
paragraph.ListFormat.ApplyStyle(numberList.Name);

// Add nested list items with different levels
paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.1");
paragraph.ListFormat.ApplyStyle(numberList.Name);
paragraph.ListFormat.ListLevelNumber = 1;

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.2");
paragraph.ListFormat.ApplyStyle(numberList.Name);
paragraph.ListFormat.ListLevelNumber = 1;

// Add deeper nested list items
paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.2.1");
paragraph.ListFormat.ApplyStyle(numberList.Name);
paragraph.ListFormat.ListLevelNumber = 2;

// Apply bulleted list style
paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 1");
paragraph.ListFormat.ApplyStyle(bulletList.Name);

paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2");
paragraph.ListFormat.ApplyStyle(bulletList.Name);

// Add nested bulleted list items
paragraph = sec.AddParagraph();
paragraph.AppendText("List Item 2.1");
paragraph.ListFormat.ApplyStyle(bulletList.Name);
paragraph.ListFormat.ListLevelNumber = 1;
```

---

# Spire.Doc JavaScript Multi-Style Paragraph
## Create a paragraph with multiple styles in a Word document
```javascript
//Create a Word document
let doc = wasmModule.Document.Create();

//Add a section
let section = doc.AddSection();

//Add a paragraph
let para = section.AddParagraph();

//Add a text range 1 and set its style
let range = para.AppendText("Spire.Doc for JavaScript ");
range.CharacterFormat.FontName = "Calibri";
range.CharacterFormat.FontSize = 16;
range.CharacterFormat.TextColor = wasmModule.Color.get_Blue();
range.CharacterFormat.Bold = true;
range.CharacterFormat.UnderlineStyle = wasmModule.UnderlineStyle.Single;

//Add a text range 2 and set its style
range = para.AppendText("is a professional Word JavaScript library");
range.CharacterFormat.FontName = "Calibri";
range.CharacterFormat.FontSize = 15;
```

---

# Spire.Doc JavaScript Paragraph Formatting
## Demonstrates how to set various paragraph formatting options in a Word document
```javascript
// Initialize a document
let document = wasmModule.Document.Create();
let sec = document.AddSection();
let para = sec.AddParagraph();
para.AppendText("Paragraph Formatting");
para.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Title});

para = sec.AddParagraph();
para.AppendText("This paragraph is surrounded with borders.");
para.Format.Borders.BorderType = wasmModule.BorderStyle.Single;
para.Format.Borders.Color = wasmModule.Color.get_Red();

para = sec.AddParagraph();
para.AppendText("The alignment of this paragraph is Left.");
para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;

para = sec.AddParagraph();
para.AppendText("The alignment of this paragraph is Center.");
para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

para = sec.AddParagraph();
para.AppendText("The alignment of this paragraph is Right.");
para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

para = sec.AddParagraph();
para.AppendText("The alignment of this paragraph is justified.");
para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Justify;

para = sec.AddParagraph();
para.AppendText("The alignment of this paragraph is distributed.");
para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Distribute;

para = sec.AddParagraph();
para.AppendText("This paragraph has the gray shadow.");
para.Format.BackColor = wasmModule.Color.get_Gray();

para = sec.AddParagraph();
para.AppendText("This paragraph has the following indentations: Left indentation is 10pt, right indentation is 10pt, first line indentation is 15pt.");
para.Format.SetLeftIndent(10);
para.Format.SetRightIndent(10);
para.Format.SetFirstLineIndent(15);

para = sec.AddParagraph();
para.AppendText("The hanging indentation of this paragraph is 15pt.");
// Negative value represents hanging indentation
para.Format.SetFirstLineIndent(-15);

para = sec.AddParagraph();
para.AppendText("This paragraph has the following spacing: spacing before is 10pt, spacing after is 20pt, line spacing is at least 10pt.");
para.Format.AfterSpacing = 20;
para.Format.BeforeSpacing = 10;
para.Format.LineSpacingRule = wasmModule.LineSpacingRule.AtLeast;
para.Format.LineSpacing = 10;
```

---

# Spire.Doc JavaScript List Management
## Restart numbered list in Word document
```javascript
//Create word document
let document = wasmModule.Document.Create();

//Create a new section
let section = document.AddSection();

//Create a new paragraph
let paragraph = section.AddParagraph();

//Append Text
paragraph.AppendText("List 1");

let numberList = wasmModule.ListStyle.Create(document, wasmModule.ListType.Numbered);
numberList.Name = "Numbered1";
document.ListStyles.Add(numberList);

//Add paragraph and apply the list style
paragraph = section.AddParagraph();
paragraph.AppendText("List Item 1");
paragraph.ListFormat.ApplyStyle(numberList.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 2");
paragraph.ListFormat.ApplyStyle(numberList.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 3");
paragraph.ListFormat.ApplyStyle(numberList.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 4");
paragraph.ListFormat.ApplyStyle(numberList.Name);

//Append Text
paragraph = section.AddParagraph();
paragraph.AppendText("List 2");

let numberList2 = wasmModule.ListStyle.Create(document, wasmModule.ListType.Numbered);
numberList2.Name = "Numbered2";
//set start number of second list
numberList2.Levels.get_Item(0).StartAt = 10;
document.ListStyles.Add(numberList2);

//Add paragraph and apply the list style
paragraph = section.AddParagraph();
paragraph.AppendText("List Item 5");
paragraph.ListFormat.ApplyStyle(numberList2.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 6");
paragraph.ListFormat.ApplyStyle(numberList2.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 7");
paragraph.ListFormat.ApplyStyle(numberList2.Name);

paragraph = section.AddParagraph();
paragraph.AppendText("List Item 8");
paragraph.ListFormat.ApplyStyle(numberList2.Name);
```

---

# Spire.Doc JavaScript Style Retrieval
## Retrieve style names from a Word document
```javascript
//Load a template document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Traverse all paragraphs in the document and get their style names through StyleName property
let styleName = "";
for (let i = 0; i < doc.Sections.Count; i++) {
    let section = doc.Sections.get(i);
    for (let j = 0; j < section.Paragraphs.Count; j++) {
        let paragraph = section.Paragraphs.get_Item(j);
        styleName += paragraph.StyleName + "\r\n";
    }
}
```

---

# Document Styles Management
## Create and apply various styles to document paragraphs
```javascript
//Initialize a document
let document = wasmModule.Document.Create();
let sec = document.AddSection();

//Add default title style to document and modify
let titleStyle = document.AddStyle(wasmModule.BuiltinStyle.Title);

titleStyle.CharacterFormat.FontName = "cambria";
titleStyle.CharacterFormat.FontSize = 28;
titleStyle.CharacterFormat.TextColor = wasmModule.Color.FromArgb(42, 123, 136);

//judge if it is Paragraph Style and then set paragraph format
if (titleStyle instanceof wasmModule.ParagraphStyle) {
    let ps = titleStyle;
    ps.ParagraphFormat.Borders.Bottom.BorderType = wasmModule.BorderStyle.Single;
    ps.ParagraphFormat.Borders.Bottom.Color = wasmModule.Color.FromArgb(42, 123, 136);
    ps.ParagraphFormat.Borders.Bottom.LineWidth = 1.5;
    ps.ParagraphFormat.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;
}

//Add default normal style and modify
let normalStyle = document.AddStyle(wasmModule.BuiltinStyle.Normal);
normalStyle.CharacterFormat.FontName = "cambria";
normalStyle.CharacterFormat.FontSize = 11;

//Add default heading1 style
let heading1Style = document.AddStyle(wasmModule.BuiltinStyle.Heading1);
heading1Style.CharacterFormat.FontName = "cambria";
heading1Style.CharacterFormat.FontSize = 14;
heading1Style.CharacterFormat.Bold = true;
heading1Style.CharacterFormat.TextColor = wasmModule.Color.FromArgb(42, 123, 136);

//Add default heading2 style
let heading2Style = document.AddStyle(wasmModule.BuiltinStyle.Heading2);
heading2Style.CharacterFormat.FontName = "cambria";
heading2Style.CharacterFormat.FontSize = 12;
heading2Style.CharacterFormat.Bold = true;

//List style
let bulletList = wasmModule.ListStyle.Create(document, wasmModule.ListType.Bulleted);
bulletList.CharacterFormat.FontName = "cambria";
bulletList.CharacterFormat.FontSize = 12;
bulletList.Name = "bulletList";
document.ListStyles.Add(bulletList);

//Apply the style
let paragraph = sec.AddParagraph();
paragraph.AppendText("Your Name");
paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Title});

paragraph = sec.AddParagraph();
paragraph.AppendText("Address, City, ST ZIP Code | Telephone | Email");
paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Normal});

paragraph = sec.AddParagraph();
paragraph.AppendText("Objective");
paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading1});

paragraph = sec.AddParagraph();
paragraph.AppendText("To get started right away, just click any placeholder text (such as this) and start typing to replace it with your own.");
paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Normal});

paragraph = sec.AddParagraph();
paragraph.AppendText("Education");
paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading1});

paragraph = sec.AddParagraph();
paragraph.AppendText("DEGREE | DATE EARNED | SCHOOL");
paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading2});

paragraph = sec.AddParagraph();
paragraph.AppendText("Major:Text");
paragraph.ListFormat.ApplyStyle("bulletList");

paragraph = sec.AddParagraph();
paragraph.AppendText("Minor:Text");
paragraph.ListFormat.ApplyStyle("bulletList");

paragraph = sec.AddParagraph();
paragraph.AppendText("Related coursework:Text");
paragraph.ListFormat.ApplyStyle("bulletList");
```

---

# Mail Merge Locale Change
## Change locale settings during mail merge operation
```javascript
// Load word document
let document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

// Store the current culture so it can be set back once mail merge is complete.
const now = new Date();
const datestr = new Intl.DateTimeFormat("de-DE",
    {year:'numeric',
        month: '2-digit',
        day:'2-digit',
        hour: "2-digit",
        minute: '2-digit',
        second: '2-digit'}).format(now);

let fieldNames = ["Contact Name", "Fax", "Date"];
let fieldValues = ["John Smith", "+1 (69) 123456", datestr];
document.MailMerge.Execute(fieldNames, fieldValues);

// Define the output file name
const outputFileName = "ChangeLocale-result.docx";

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

// Clean up resources
document.Dispose();
```

---

# Document Conditional Field Execution
## Create and execute conditional IF fields with merge fields in a document
```javascript
let doc = wasmModule.Document.Create();
//Add a new section
let section = doc.AddSection();
//Add a new paragraph for a section
let paragraph = section.AddParagraph();

CreateIFField1(doc, paragraph);
paragraph = section.AddParagraph();
CreateIFField2(doc, paragraph);

let fieldName = ["Count", "Age"];
let fieldValue = ["2", "30"];

doc.MailMerge.Execute(fieldName, fieldValue);
doc.IsUpdateFields = true;

function CreateIFField1( document,  paragraph) {
  let ifField = wasmModule.IfField.Create(document);
  ifField.Type = wasmModule.FieldType.FieldIf;
  ifField.Code = "IF ";
  paragraph.Items.Add(ifField);

  paragraph.AppendField("Count", wasmModule.FieldType.FieldMergeField);
  paragraph.AppendText(" > ");
  paragraph.AppendText("\"1\" ");
  paragraph.AppendText("\"Greater than one\" ");
  paragraph.AppendText("\"Less than one\"");

  let end = document.CreateParagraphItem(wasmModule.ParagraphItemType.FieldMark);
  end.Type = wasmModule.FieldMarkType.FieldEnd;
  paragraph.Items.Add(end);

  ifField.End = end;
}

function CreateIFField2( document,  paragraph) {
  let ifField = wasmModule.IfField.Create(document);
  ifField.Type = wasmModule.FieldType.FieldIf;
  ifField.Code = "IF ";
  paragraph.Items.Add(ifField);

  paragraph.AppendField("Age", wasmModule.FieldType.FieldMergeField);
  paragraph.AppendText(" > ");
  paragraph.AppendText("\"50\" ");
  paragraph.AppendText("\"The old man\" ");
  paragraph.AppendText("\"The young man\"");

  let end = document.CreateParagraphItem(wasmModule.ParagraphItemType.FieldMark);
  end.Type = wasmModule.FieldMarkType.FieldEnd;
  paragraph.Items.Add(end);

  ifField.End = end;
}
```

---

# spire.doc javascript mail merge
## hide empty regions during mail merge
```javascript
//Create word document
let document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);
let filedNames = ["Contact Name", "Fax", "Date"];
let filedValues = ["John Smith", "+1 (69) 123456", wasmModule.DateTime.get_Now().Date.ToString()];
//Set the value to remove paragraphs which contain empty field.
document.MailMerge.HideEmptyParagraphs = true;
//Set the value to remove group which contain empty field.
document.MailMerge.HideEmptyGroup = true;
document.MailMerge.Execute(filedNames, filedValues);
```

---

# Spire.Doc JavaScript Mail Merge
## Identify merge field names in Word document
```javascript
//Create Word document.
let document = wasmModule.Document.Create();

//Load the file from disk.
document.LoadFromFile(inputFileName);

//Get the collection of group names.
let GroupNames = document.MailMerge.GetMergeGroupNames();

//Get the collection of merge field names in a specific group.
let MergeFieldNamesWithinRegion = document.MailMerge.GetMergeFieldNames({groupName: "Products"});

//Get the collection of all the merge field names.
let MergeFieldNames = document.MailMerge.GetMergeFieldNames();

document.Close();
```

---

# Spire.Doc JavaScript Mail Merge
## Demonstrates how to merge mail into a Word document
```javascript
// Define input file name
let inputFileName = "MailMerage.docx";

// Define output file name
const outputFileName = "MailMerage-result.docx";

// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Define field names for mail merge
let filedNames = ["Contact Name", "Fax", "Date"];

// Define field values for mail merge
let filedValues = ["John Smith", "+1 (69) 123456", wasmModule.DateTime.get_Now().Date.ToString()];

// Execute mail merge with field names and values
document.MailMerge.Execute(filedNames, filedValues);

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013});

// Clean up resources
document.Dispose();
```

---

# spire.doc javascript mail merge
## execute mail merge with switches
```javascript
// Create a new document
const doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

let fieldName = ["XX_Name"];
let fieldValue = ["Jason Tang"];

doc.MailMerge.Execute(fieldName, fieldValue);
```

---

# Spire.Doc JavaScript Bookmark
## Copy bookmark content in a Word document
```javascript
//Get the bookmark by name.
let bookmark = doc.Bookmarks._get_Item("Test");
let docObj = null;

//Judge if the paragraph includes the bookmark exists in the table, if it exists in cell,
//Then need to find its outermost parent object(Table),
//and get the start/end index of current object on body.
if (bookmark.BookmarkStart.Owner.IsInCell) {
    docObj = bookmark.BookmarkStart.Owner.Owner.Owner.Owner;
} else {
    docObj = bookmark.BookmarkStart.Owner;
}
let startIndex = doc.Sections.get(0).Body.ChildObjects.IndexOf(docObj);
if (bookmark.BookmarkEnd.Owner.IsInCell) {
    docObj = bookmark.BookmarkEnd.Owner.Owner.Owner.Owner;
} else {
    docObj = bookmark.BookmarkEnd.Owner;
}
let endIndex = doc.Sections.get(0).Body.ChildObjects.IndexOf(docObj);

//Get the start/end index of the bookmark object on the paragraph.
let para = bookmark.BookmarkStart.Owner;
let pStartIndex = para.ChildObjects.IndexOf(bookmark.BookmarkStart);
para = bookmark.BookmarkEnd.Owner;
let pEndIndex = para.ChildObjects.IndexOf(bookmark.BookmarkEnd);

//Get the content of current bookmark and copy.
let select = wasmModule.TextBodySelection.Create(doc.Sections.get_Item(0).Body, startIndex, endIndex, pStartIndex, pEndIndex);
let body = wasmModule.TextBodyPart.CreateByTextBodySelection(select);
for (let i = 0; i < body.BodyItems.Count; i++) {
    doc.Sections.get(0).Body.ChildObjects.Add(body.BodyItems.get(i).Clone());
}
```

---

# Word Document Bookmark Creation
## Create simple and nested bookmarks in a Word document
```javascript
function CreateBookmarkBase(section) {
  let paragraph = section.AddParagraph();
  let txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark in a Word document.");
  txtRange.CharacterFormat.Italic = true;

  section.AddParagraph();
  paragraph = section.AddParagraph();
  txtRange = paragraph.AppendText("Simple Create Bookmark.");
  txtRange.CharacterFormat.TextColor = wasmModule.Color.get_CornflowerBlue();
  paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading2});

  //Write simple CreateBookmarks.
  section.AddParagraph();
  paragraph = section.AddParagraph();
  paragraph.AppendBookmarkStart("SimpleCreateBookmark");
  paragraph.AppendText("This is a simple bookmark.");
  paragraph.AppendBookmarkEnd("SimpleCreateBookmark");

  section.AddParagraph();
  paragraph = section.AddParagraph();
  txtRange = paragraph.AppendText("Nested Create Bookmark.");
  txtRange.CharacterFormat.TextColor = wasmModule.Color.get_CornflowerBlue();
  paragraph.ApplyStyle({builtinStyle : wasmModule.BuiltinStyle.Heading2});

  //Write nested CreateBookmarks.
  section.AddParagraph();
  paragraph = section.AddParagraph();
  paragraph.AppendBookmarkStart("Root");
  txtRange = paragraph.AppendText(" This is Root data ");
  txtRange.CharacterFormat.Italic = true;
  paragraph.AppendBookmarkStart("NestedLevel1");
  txtRange = paragraph.AppendText(" This is Nested Level1 ");
  txtRange.CharacterFormat.Italic = true;
  txtRange.CharacterFormat.TextColor = wasmModule.Color.get_DarkSlateGray();
  paragraph.AppendBookmarkStart("NestedLevel2");
  txtRange = paragraph.AppendText(" This is Nested Level2 ");
  txtRange.CharacterFormat.Italic = true;
  txtRange.CharacterFormat.TextColor = wasmModule.Color.get_DimGray();
  paragraph.AppendBookmarkEnd("NestedLevel2");
  paragraph.AppendBookmarkEnd("NestedLevel1");
  paragraph.AppendBookmarkEnd("Root");
}
```

---

# Spire.Doc JavaScript Bookmark
## Create bookmark for table in Word document
```javascript
// Create bookmark for a table
function CreateBookmarkForTableBase(doc, section) {
  // Add a paragraph
  let paragraph = section.AddParagraph();

  // Append text for added paragraph
  let txtRange = paragraph.AppendText("The following example demonstrates how to create bookmark for a table in a Word document.");

  // Set the font in italic
  txtRange.CharacterFormat.Italic = true;

  // Append bookmark start
  paragraph.AppendBookmarkStart("CreateBookmark");

  // Append bookmark end
  paragraph.AppendBookmarkEnd("CreateBookmark");

  // Add table
  let table = section.AddTable({showBorder: true});

  // Set the number of rows and columns
  table.ResetCells(2, 2);

  // Append text for table cells
  let range = table.Rows.get(0).Cells.get(0).AddParagraph().AppendText("sampleA");
  range = table.Rows.get(0).Cells.get(1).AddParagraph().AppendText("sampleB");
  range = table.Rows.get(1).Cells.get(0).AddParagraph().AppendText("120");
  range = table.Rows.get(1).Cells.get(1).AddParagraph().AppendText("260");

  // Get the bookmark by index
  let bookmark = doc.Bookmarks._get_ItemI(0);

  // Get the name of bookmark
  let bookmarkName = bookmark.Name;

  // Locate the bookmark by name
  let navigator = wasmModule.BookmarksNavigator.Create(doc);
  navigator.MoveToBookmark(bookmarkName);

  // Add table to TextBodyPart
  let part = navigator.GetBookmarkContent();
  part.BodyItems.Add(table);

  // Replace bookmark content with table
  navigator.ReplaceBookmarkContent({bodyPart : part});
}
```

---

# Extract Bookmark Text
## Extract text content from a bookmark in a Word document
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Creates a BookmarkNavigator instance to access the bookmark
let navigator = wasmModule.BookmarksNavigator.Create(doc);
// Locate a specific bookmark by bookmark name
navigator.MoveToBookmark("Content");
let textBodyPart = navigator.GetBookmarkContent();

// Iterate through the items in the bookmark content to get the text
let text = "";
for (let i = 0; i < textBodyPart.BodyItems.Count; i++) {
    let item = textBodyPart.BodyItems.get(i);
    if (item instanceof wasmModule.Paragraph) {
        for (let j = 0; j < item.ChildObjects.Count; j++) {
            let childObject = item.ChildObjects.get(j);
            if (childObject instanceof wasmModule.TextRange) {
                text += childObject.Text;
            }
        }
    }
}
doc.Close();
```

---

# spire.doc javascript bookmarks
## get bookmarks by index and name from word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

//Get the bookmark by index.
let bookmark1 = document.Bookmarks._get_ItemI(0);

//Get the bookmark by name.
let bookmark2 = document.Bookmarks._get_Item("Test2");

//Create content array to save
let content = [];

//Set string format for displaying
let result = "The bookmark obtained by index is " + bookmark1.Name + ".\r\nThe bookmark obtained by name is " + bookmark2.Name + ".\n";

//Add result string to content array
content.push(result);

// Define the output file name
const outputFileName = "GetBookmarks-result.txt";

//Write the contents in a TXT file
wasmModule.FS.writeFile(outputFileName, content.join("\n"));
document.Close();
```

---

# Spire.Doc JavaScript Bookmark Insertion
## Insert a document at the location of a bookmark in a Word document
```javascript
//Create the first document
let document1 = wasmModule.Document.Create();

//Load the first document from disk.
document1.LoadFromFile(inputFileName_1);

//Create the second document
let document2 = wasmModule.Document.Create();

//Load the second document from disk.
document2.LoadFromFile(inputFileName_2);

//Get the first section of the first document
let section1 = document1.Sections.get_Item();

//Locate the bookmark
let bn = wasmModule.BookmarksNavigator.Create(document1);

//Find bookmark by name
bn.MoveToBookmark("Test", true, true);

//Get bookmarkStart
let start = bn.CurrentBookmark.BookmarkStart;

//Get the owner paragraph
let para = start.OwnerParagraph;

//Get the para index
let index = section1.Body.ChildObjects.IndexOf(para);

//Insert the paragraphs of document2
for (let i = 0; i < document2.Sections.Count; i++) {
    let section2 = document2.Sections.get_Item(i);
    for (let j = 0; j < section2.Paragraphs.Count; j++) {
        let paragraph = section2.Paragraphs.get_Item(j);
        section1.Body.ChildObjects.Insert(index + 1, paragraph.Clone());
    }
}
```

---

# Insert Image at Bookmark in Word Document
## This code demonstrates how to insert an image at the location of a bookmark in a Word document
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

//Load the document
doc.LoadFromFile(inputFileName);

//Create an instance of BookmarksNavigator
let bn = wasmModule.BookmarksNavigator.Create(doc);

//Find a bookmark named Test
bn.MoveToBookmark("Test", true, true);

//Add a section
let section0 = doc.AddSection();

//Add a paragraph for the section
let paragraph = section0.AddParagraph();

//Add a picture into the paragraph
paragraph.AppendPicture({imgFile : imageFileName});

//Add the paragraph at the position of bookmark
bn.InsertParagraph(paragraph);

//Remove the section0
doc.Sections.Remove(section0);
```

---

# Spire.Doc JavaScript Bookmark
## Remove bookmark from Word document
```javascript
//Get the bookmark by name.
let bookmark = document.Bookmarks.get_Item("Test");

//Remove the bookmark, not its content.
document.Bookmarks.Remove(bookmark);
```

---

# spire.doc javascript bookmark
## remove bookmark content in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load the document from disk
document.LoadFromFile(inputFileName);

// Get the bookmark by name
let bookmark = document.Bookmarks._get_Item("Test");

let para = bookmark.BookmarkStart.Owner;
let startIndex = para.ChildObjects.IndexOf(bookmark.BookmarkStart);
para = bookmark.BookmarkEnd.Owner;
let endIndex = para.ChildObjects.IndexOf(bookmark.BookmarkEnd);

// Remove the content object, and Start from next of BookmarkStart object, end up with previous of BookmarkEnd object.
// This method is only to remove the content of the bookmark.
for (let i = startIndex + 1; i < endIndex; i++) {
    para.ChildObjects.RemoveAt(startIndex + 1);
}
```

---

# Word Document Bookmark Replacement
## Replace bookmark content in a Word document
```javascript
// Locate the bookmark
let bookmarkNavigator = wasmModule.BookmarksNavigator.Create(doc);
bookmarkNavigator.MoveToBookmark("Test");

// Replace the context with new
bookmarkNavigator.ReplaceBookmarkContent({text: "This is replaced content.", saveFormatting: false});
```

---

# Spire.Doc JavaScript Bookmark Replacement
## Replace bookmark content with a table in Word document
```javascript
//Create a table
let table = wasmModule.Table.Create(doc, true);
table.ResetCells(4,5);
//Create datatable
let dt = [       //not supported datatable
    //["id", "name", "job", "email", "salary"],
    ["Name", "Capital", "Continent", "Area", "Population"],
    ["Argentina", "Buenos Aires", "South America", "2777815", "32300003"],
    ["Bolivia", "La Paz", "South America", "1098575", "7300000"],
    ["Brazil", "Brasilia", "South America", "8511196", "150400000"]];
//Fill the table with the data of datatable
for (let i = 0; i < 4; i++) {
    for (let j = 0; j < 5; j++) {
        table.Rows.get(i).Cells.get(j).AddParagraph().AppendText(dt[i][j]);
      }
}

//Get the specific bookmark by its name
let navigator = wasmModule.BookmarksNavigator.Create(doc);
navigator.MoveToBookmark("Test");

//Create a TextBodyPart instance and add the table to it
let part = wasmModule.TextBodyPart.Create(doc);
part.BodyItems.Add(table);

//Replace the current bookmark content with the TextBodyPart object
navigator.ReplaceBookmarkContent({bodyPart : part});
```

---

# Spire.Doc JavaScript Comment
## Add comment for specific text in a document
```javascript
function InsertComments(doc, keystring) {
  //Find the key string
  let find = doc.FindString(keystring, false, true);

  //Create the commentmarkStart and commentmarkEnd
  let commentmarkStart = wasmModule.CommentMark.Create(doc);
  commentmarkStart.Type = wasmModule.CommentMarkType.CommentStart;
  let commentmarkEnd = wasmModule.CommentMark.Create(doc);
  commentmarkEnd.Type = wasmModule.CommentMarkType.CommentEnd;

  //Add the content for comment
  let comment = wasmModule.Comment.Create(doc);
  comment.Body.AddParagraph().Text = "Test comments";
  comment.Format.Author = "E-iceblue";

  //Get the textRange
  let range = find.GetAsOneRange();

  //Get its paragraph
  let para = range.OwnerParagraph;

  //Get the index of textRange
  let index = para.ChildObjects.IndexOf(range);

  //Add comment
  para.ChildObjects.Add(comment);

  //Insert the commentmarkStart and commentmarkEnd
  para.ChildObjects.Insert(index, commentmarkStart);
  para.ChildObjects.Insert(index + 2, commentmarkEnd);
}
```

---

# Spire.Doc JavaScript Comment
## Insert comment in Word document
```javascript
function InsertComments(section) {
  //Insert comment.
  let paragraph = section.Paragraphs.get_Item(1);
  let comment = paragraph.AppendComment("Spire.Doc for .NET");
  comment.Format.Author = "E-iceblue";
  comment.Format.Initial = "CM";
}
```

---

# Extract Comments from Word Document
## This example demonstrates how to extract all comments from a Word document and save their text content

```javascript
// Create a new document object  
let doc = wasmModule.Document.Create();
// Load file from VFS
doc.LoadFromFile(inputFileName);

// Initialize an empty array to store the extracted text from the comments
let SB = [];

//Traverse all comments
for (let i = 0; i < doc.Comments.Count; i++) {
  let comment = doc.Comments.get_Item(i);
  for (let j = 0; j < comment.Body.Paragraphs.Count; j++) {
    let p = comment.Body.Paragraphs.get_Item(j);
    SB.push(p.Text + "\n");
  }
}
```

---

# spire.doc javascript comment
## insert picture into comment
```javascript
// Create a new document object   
let doc = wasmModule.Document.Create();
// Load file from VFS
doc.LoadFromFile(inputFileName);

// Get the first paragraph and insert comment
let paragraph = doc.Sections.get(0).Paragraphs.get_Item(2);
let comment = paragraph.AppendComment("This is a comment.");
comment.Format.Author = "E-iceblue";

// Load a picture
let docPicture = wasmModule.DocPicture.Create(doc);
docPicture.LoadImage(imageFile);

// Insert the picture into the comment body
comment.Body.AddParagraph().ChildObjects.Add(docPicture);
```

---

# Spire.Doc JavaScript Comment Operations
## Remove and replace comments in a Word document
```javascript
// Replace the content of the first comment
doc.Comments.get_Item(0).Body.Paragraphs.get_Item(0).Replace({ given: "This is the title", replace: "This comment is changed.", caseSensitive: false, wholeWord: false });

// Remove the second comment
doc.Comments.RemoveAt(1);
```

---

# Spire.Doc JavaScript Comment Processing
## Remove content with comment in Word document
```javascript
// Create a new document object  
let document = wasmModule.Document.Create();

// Load the document from VFS
document.LoadFromFile(inputFileName);

// Get the first comment
let comment = document.Comments.get_Item(0);

// Get the paragraph of obtained comment
let para = comment.OwnerParagraph;

// Get index of the CommentMarkStart
let startIndex = para.ChildObjects.IndexOf(comment.CommentMarkStart);

// Get index of the CommentMarkEnd
let endIndex = para.ChildObjects.IndexOf(comment.CommentMarkEnd);

// Create a list
let list = [];

// Get TextRanges between the indexes
for (let i = startIndex; i < endIndex; i++) {
  if (para.ChildObjects.get(i) instanceof wasmModule.TextRange) {
    list.push(para.ChildObjects.get(i));
  }
}

// Insert a new TextRange
let textRange = wasmModule.TextRange.Create(document);

// Set text is null
textRange.Text = "";

// Insert the new textRange
para.ChildObjects.Insert(endIndex, textRange);

// Remove previous TextRanges
for (let i = 0; i < list.length; i++) {
  para.ChildObjects.Remove(list[i]);
}
```

---

# Spire.Doc JavaScript Comment
## Reply to comment in Word document
```javascript
// Create a new document object  
let doc = wasmModule.Document.Create();

// Load the document
doc.LoadFromFile(inputFileName);

// Get the first comment
let comment1 = doc.Comments.get_Item(0);

// Create a new comment and specify the author and content
let replyComment1 = wasmModule.Comment.Create(doc);
replyComment1.Format.Author = "E-iceblue";
replyComment1.Body.AddParagraph().AppendText("Spire.Doc is a professional Word .NET library on operating Word documents.");

// Add the new comment as a reply to the selected comment
comment1.ReplyToComment(replyComment1);

// Load image
let docPicture = wasmModule.DocPicture.Create(doc);
docPicture.LoadImage(imageFile);

// Insert a picture in the comment
replyComment1.Body.Paragraphs.get_Item(0).ChildObjects.Add(docPicture);
```

---

# Spire.Doc JavaScript Barcode Image
## Add barcode image to Word document
```javascript
// Create a new document object   
let document = wasmModule.Document.Create();

// Load file from VFS
document.LoadFromFile(inputFileName);

// Add barcode image
document.Sections.get(0).AddParagraph().AppendPicture({ imgFile: imageFile });
```

---

# spire.doc javascript horizontal line
## add horizontal line to word document
```javascript
// Create a new document object  
let doc = wasmModule.Document.Create();

// Add a section
let sec = doc.AddSection();

// Add a paragraph
let para = sec.AddParagraph();

// Add HorizonalLine
para.AppendHorizonalLine();
```

---

# Spire.Doc JavaScript Image and Shape
## Add image and textbox to each page in a Word document
```javascript
// Create a new document object   
let document = wasmModule.Document.Create();

// Add a picture in footer and set its position
let picture = document.Sections.get(0).HeadersFooters.Footer.AddParagraph().AppendPicture({ imgFile: imgPathName });
picture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
picture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
picture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Bottom;
picture.TextWrappingStyle = wasmModule.TextWrappingStyle.None;

// Add a textbox in footer and set its position
let textbox = document.Sections.get(0).HeadersFooters.Footer.AddParagraph().AppendTextBox(150, 20);
textbox.VerticalOrigin = wasmModule.VerticalOrigin.Page;
textbox.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
textbox.HorizontalPosition = 300;
textbox.VerticalPosition = 700;
textbox.Body.AddParagraph().AppendText("Welcome to E-iceblue");
```

---

# Adding Shape Groups in Word Documents
## Demonstrates how to create and add shape groups with various shapes to a Word document
```javascript
// Create a new document object  
let doc = wasmModule.Document.Create();
let sec = doc.AddSection();

// Add a new paragraph
let para = sec.AddParagraph();
// Add a shape group with the height and width
let shapegroup = para.AppendShapeGroup(375, 462);
shapegroup.HorizontalPosition = 180;

// Calculate the scale ratio
let X = (shapegroup.Width / 1000.0);
let Y = (shapegroup.Height / 1000.0);

// Create a textbox
let txtBox = wasmModule.TextBox.Create(doc);
txtBox.SetShapeType(wasmModule.ShapeType.RoundRectangle);
txtBox.Width = 125 / X;
txtBox.Height = 54 / Y;
let paragraph = txtBox.Body.AddParagraph();
paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
paragraph.AppendText("Start");
txtBox.HorizontalPosition = 19 / X;
txtBox.VerticalPosition = 27 / Y;
txtBox.Format.LineColor = wasmModule.Color.get_Green();
shapegroup.ChildObjects.Add(txtBox);

// Create an arrow line shape
let arrowLineShape = wasmModule.ShapeObject.Create(doc, wasmModule.ShapeType.DownArrow);
arrowLineShape.Width = 16 / X;
arrowLineShape.Height = 40 / Y;
arrowLineShape.HorizontalPosition = 69 / X;
arrowLineShape.VerticalPosition = 87 / Y;
arrowLineShape.StrokeColor = wasmModule.Color.get_Purple();
shapegroup.ChildObjects.Add(arrowLineShape);

// Create a textbox
txtBox = wasmModule.TextBox.Create(doc);
txtBox.SetShapeType(wasmModule.ShapeType.Rectangle);
txtBox.Width = 125 / X;
txtBox.Height = 54 / Y;
paragraph = txtBox.Body.AddParagraph();
paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
paragraph.AppendText("Step 1");
txtBox.HorizontalPosition = 19 / X;
txtBox.VerticalPosition = 131 / Y;
txtBox.Format.LineColor = wasmModule.Color.get_Blue();
shapegroup.ChildObjects.Add(txtBox);

// Create an arrow line shape
arrowLineShape = wasmModule.ShapeObject.Create(doc, wasmModule.ShapeType.DownArrow);
arrowLineShape.Width = 16 / X;
arrowLineShape.Height = 40 / Y;
arrowLineShape.HorizontalPosition = 69 / X;
arrowLineShape.VerticalPosition = 192 / Y;
arrowLineShape.StrokeColor = wasmModule.Color.get_Purple();
shapegroup.ChildObjects.Add(arrowLineShape);

// Create an arrow line shape
txtBox = wasmModule.TextBox.Create(doc);
txtBox.SetShapeType(wasmModule.ShapeType.Parallelogram);
txtBox.Width = 149 / X;
txtBox.Height = 59 / Y;
paragraph = txtBox.Body.AddParagraph();
paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
paragraph.AppendText("Step 2");
txtBox.HorizontalPosition = 7 / X;
txtBox.VerticalPosition = 236 / Y;
txtBox.Format.LineColor = wasmModule.Color.get_BlueViolet();
shapegroup.ChildObjects.Add(txtBox);

// Create an arrow line shape
arrowLineShape = wasmModule.ShapeObject.Create(doc, wasmModule.ShapeType.DownArrow);
arrowLineShape.Width = 16 / X;
arrowLineShape.Height = 40 / Y;
arrowLineShape.HorizontalPosition = 66 / X;
arrowLineShape.VerticalPosition = 300 / Y;
arrowLineShape.StrokeColor = wasmModule.Color.get_Purple();
shapegroup.ChildObjects.Add(arrowLineShape);

// Create an arrow line shape
txtBox = wasmModule.TextBox.Create(doc);
txtBox.SetShapeType(wasmModule.ShapeType.Rectangle);
txtBox.Width = 125 / X;
txtBox.Height = 54 / Y;
paragraph = txtBox.Body.AddParagraph();
paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
paragraph.AppendText("Step 3");
txtBox.HorizontalPosition = 19 / X;
txtBox.VerticalPosition = 345 / Y;
txtBox.Format.LineColor = wasmModule.Color.get_Blue();
shapegroup.ChildObjects.Add(txtBox);
```

---

# spire.doc javascript shapes
## add various shapes to word document
```javascript
// Create a new document
let doc = wasmModule.Document.Create();

// Add a section
let sec = doc.AddSection();

// Add a paragraph
let para = sec.AddParagraph();
let x = 60, y = 40, lineCount = 0;
for (let i = 1; i < 20; i++) {
  if (lineCount > 0 && lineCount % 8 == 0) {
    para.AppendBreak(wasmModule.BreakType.PageBreak);
    x = 60;
    y = 40;
    lineCount = 0;
  }
  //Add shape and set its size and position.
  let shape = para.AppendShape(50, 50, wasmModule.ShapeType.fromValue(i));
  shape.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
  shape.HorizontalPosition = x;
  shape.VerticalOrigin = wasmModule.VerticalOrigin.Page;
  shape.VerticalPosition = y + 50;
  x = x + shape.Width + 50;
  if (i > 0 && i % 5 == 0) {
    y = y + shape.Height + 120;
    lineCount++;
    x = 60;
  }
}
```

---

# Spire.Doc JavaScript Shape Alignment
## Align shapes horizontally in a Word document
```javascript
// Get first section
let section = doc.Sections.get_Item(0);

for (let i = 0; i < section.Paragraphs.Count; i++) {
  let para = section.Paragraphs.get_Item(i);
  for (let j = 0; j < para.ChildObjects.Count; j++) {
    let obj = para.ChildObjects.get(j);
    if (obj instanceof wasmModule.ShapeObject) {
      //Set the horizontal alignment as center
      obj.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Center;

      ////Set the vertical alignment as top
      //(obj as ShapeObject).VerticalAlignment = ShapeVerticalAlignment.Top;
    }
  }
}
```

---

# Get Alternative Text from Word Document Shapes
## This code demonstrates how to retrieve alternative text from shapes in a Word document.
```javascript
// Create a document
let document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);

// Create string builder
let builder = [];

// Loop through shapes and get the AlternativeText
for (let i = 0; i < document.Sections.Count; i++) {
  let section = document.Sections.get(i);
  for (let j = 0; j < section.Paragraphs.Count; j++) {
    let para = section.Paragraphs.get_Item(j);
    for (let k = 0; k < para.ChildObjects.Count; k++) {
      let obj = para.ChildObjects.get(k);
      if (obj instanceof wasmModule.ShapeObject) {
        let text = obj.AlternativeText;
        // Append the alternative text in builder
        builder.push(text + "\n");
      }
    }
  }
}

// Dispose of the document object to free resources
document.Dispose();
```

---

# spire.doc javascript image
## insert image into word document
```javascript
function InsertImage(section, imageFile) {
  // Add paragraph
  let paragraph = section.AddParagraph();
  paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;

  let picture = paragraph.AppendPicture({ imgFile: imageFile });

  picture.Width = 100;
  picture.Height = 100;

  paragraph = section.AddParagraph();
  paragraph.Format.LineSpacing = 20;
  let tr = paragraph.AppendText("Spire.Doc for .NET is a professional Word .NET library specially designed for developers to create, read, write, convert and print Word document files from any .NET( C#, VB.NET, ASP.NET) platform with fast and high quality performance. ");
  tr.CharacterFormat.FontName = "Arial";
  tr.CharacterFormat.FontSize = 14;

  section.AddParagraph();
  paragraph = section.AddParagraph();
  paragraph.Format.LineSpacing = 20;
  tr = paragraph.AppendText("As an independent Word .NET component, Spire.Doc for .NET doesn't need Microsoft Word to be installed on the machine. However, it can incorporate Microsoft Word document creation capabilities into any developers' .NET applications.");
  tr.CharacterFormat.FontName = "Arial";
  tr.CharacterFormat.FontSize = 14;
}
```

---

# spire.doc javascript image insertion
## insert image into word document
```javascript
// Create a picture
let picture = wasmModule.DocPicture.Create(doc);
picture.LoadImage({ imgFile: inputImgFileName });
// Set image's position
picture.HorizontalPosition = 50.0;
picture.VerticalPosition = 60.0;

// Set image's size
picture.Width = 200;
picture.Height = 200;

// Set textWrappingStyle with image;
picture.TextWrappingStyle = wasmModule.TextWrappingStyle.Through;
// Insert the picture at the beginning of the second paragraph
paragraph.ChildObjects.Insert(0, picture);
```

---

# spire.doc javascript wordart
## insert WordArt in a Word document
```javascript
// Create Word document
let doc = wasmModule.Document.Create();

// Load Word document
doc.LoadFromFile(inputFileName);

// Add a paragraph
let paragraph = doc.Sections.get(0).AddParagraph();

// Add a shape
let shape = paragraph.AppendShape(250, 70, wasmModule.ShapeType.TextWave4);

// Set the position of the shape
shape.VerticalPosition = 20;
shape.HorizontalPosition = 80;

// Set the text of WordArt
shape.WordArt.Text = "Thanks for reading.";

// Set the fill color
shape.FillColor = wasmModule.Color.get_Red();

// Set the border color of the text
shape.StrokeColor = wasmModule.Color.get_Yellow();
```

---

# spire.doc javascript shape removal
## remove shapes from word document
```javascript
// Load Document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);
let section = doc.Sections.get(0);
//Get all the child objects of paragraph
for (let i = 0; i < section.Paragraphs.Count; i++) {
  let para = section.Paragraphs.get_Item(i);
  for (let j = 0; j < para.ChildObjects.Count; j++) {
    //If the child objects is shape object
    if (para.ChildObjects.get(j) instanceof wasmModule.ShapeObject) {
      //Remove the shape object
      para.ChildObjects.RemoveAt(j);
      j--;
    }
  }
}
```

---

# Spire.Doc JavaScript Image Replacement
## Replace images with text in a Word document
```javascript
// Load Document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

// Replace all pictures with texts
let a = 1;
for (let i = 0; i < doc.Sections.Count; i++) {
  let sec = doc.Sections.get(i);
  for (let j = 0; j < sec.Paragraphs.Count; j++) {
    let para = sec.Paragraphs.get_Item(j);
    let pictures = [];
    //Get all pictures in the Word document
    for (let k = 0; k < para.ChildObjects.Count; k++) {
      let docObj = para.ChildObjects.get(k);
      if (docObj.DocumentObjectType == wasmModule.DocumentObjectType.Picture) {
        pictures.push(docObj);
      }
    }

    //Replace pictures with the text "Here was image {image index}"
    for (let i = 0; i < pictures.length; i++) {
      let pic = pictures[i];
      let index = para.ChildObjects.IndexOf(pic);
      let range = wasmModule.TextRange.Create(doc);
      range.Text = "Here was image " + a
      para.ChildObjects.Insert(index, range);
      para.ChildObjects.Remove(pic);
      a++;
    }
  }
}
```

---

# Reset Image Size in Word Document
## This example demonstrates how to reset the size of images in a Word document
```javascript
// Load Document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

// Get the first secion
let section = doc.Sections.get(0);
// Get the first paragraph
let paragraph = section.Paragraphs.get_Item(0);

// Reset the image size of the first paragraph
for (let i = 0; i < paragraph.ChildObjects.Count; i++) {
  let docObj = paragraph.ChildObjects.get(i);
  if (docObj instanceof wasmModule.DocPicture) {
    let picture = docObj;
    picture.Width = 50;
    picture.Height = 50;
  }
}
```

---

# spire.doc javascript shape
## reset shape size in word document
```javascript
// Get the first section and the first paragraph that contains the shape
let section = doc.Sections.get(0);
let para = section.Paragraphs.get_Item(0);

// Get the second shape and reset the width and height for the shape
let shape = para.ChildObjects.get(1);
shape.Width = 200;
shape.Height = 200;
```

---

# spire.doc javascript shape rotation
## rotate shapes in word document
```javascript
// Get the first section
let section = doc.Sections.get(0);

// Traverse the word document and set the shape rotation as 20
for (let i = 0; i < section.Paragraphs.Count; i++) {
  let para = section.Paragraphs.get_Item(i);
  for (let j = 0; j < para.ChildObjects.Count; j++) {
    let obj = para.ChildObjects.get(j);
    if (obj instanceof wasmModule.ShapeObject) {
      obj.Rotation = 20.0;
    }
  }
}
```

---

# Spire.Doc JavaScript Line Shape Style
## Set style properties for line shape in Word document
```javascript
// Create a document
let doc = wasmModule.Document.Create();

// Add a section
let sec = doc.AddSection();

// Add a new paragraph
let para = sec.AddParagraph();

// Add a line shape
let shape = para.AppendShape(100, 100, wasmModule.ShapeType.Line);

// Set style of Line shape
shape.FillColor = wasmModule.Color.get_Orange();
shape.StrokeColor = wasmModule.Color.get_Black();
shape.LineStyle = wasmModule.ShapeLineStyle.Single;
shape.LineDashing = wasmModule.LineDashing.LongDashDotDotGEL;
```

---

# Spire.Doc JavaScript Text Wrap
## Set text wrap style for images in a Word document
```javascript
// Iterate through all sections in the document
for (let i = 0; i < doc.Sections.Count; i++) {
  let sec = doc.Sections.get(i);
  for (let j = 0; j < sec.Paragraphs.Count; j++) {
    let pictures = [];
    let para = sec.Paragraphs.get_Item(j);
    // Get all pictures in the Word document
    for (let k = 0; k < para.ChildObjects.Count; k++) {
      let docObj = para.ChildObjects.get(k);
      if (docObj.DocumentObjectType == wasmModule.DocumentObjectType.Picture) {
        pictures.push(docObj);
      }
    }
    // Set text wrap styles for each picture
    for (let pic of pictures) {
      let picture = pic;
      picture.TextWrappingStyle = wasmModule.TextWrappingStyle.Through;
      picture.TextWrappingType = wasmModule.TextWrappingType.Both;
    }
  }
}
```

---

# Word Document Textbox Transparency
## Set transparency for textbox in Word document
```javascript
// Create a word document
let doc = wasmModule.Document.Create();

// Create a new section
let section = doc.AddSection();

// Create a new paragraph
let paragraph = section.AddParagraph();

// Append TextBox
let textbox1 = paragraph.AppendTextBox(100, 50);

// Set fill color
textbox1.Format.FillColor = wasmModule.Color.get_Red();

// Set fill transparency
textbox1.FillTransparency = 0.45;
```

---

# spire.doc javascript image transparency
## set transparent color for images in word document
```javascript
// Get the first paragraph in the first section
let paragraph = doc.Sections.get(0).Paragraphs.get_Item(0);

// Set the blue color of the image(s) in the paragraph to transparent
for (let i = 0; i < paragraph.ChildObjects.Count; i++) {
  let obj = paragraph.ChildObjects.get(i);
  if (obj instanceof wasmModule.DocPicture) {
    let picture = obj;
    picture.TransparentColor = wasmModule.Color.get_Blue();
  }
}
```

---

# Spire.Doc JavaScript Image Update
## Replace an image in a Word document with a new image
```javascript
// Load Document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

// Get all pictures in the Word document
let pictures = []
for (let i = 0; i < doc.Sections.Count; i++) {
  let sec = doc.Sections.get(i);
  for (let j = 0; j < sec.Paragraphs.Count; j++) {
    let para = sec.Paragraphs.get_Item(j);
    for (let k = 0; k < para.ChildObjects.Count; k++) {
      let docObj = para.ChildObjects.get(k);
      if (docObj.DocumentObjectType == wasmModule.DocumentObjectType.Picture) {
        pictures.push(docObj);
      }
    }
  }
}

// Replace the first picture with a new image file
let picture = pictures[0]
picture.LoadImage({ imgFile: imageFileName });

// Save the document
doc.SaveToFile({ fileName: "UpdateImage.docx", fileFormat: wasmModule.FileFormat.Docx });

// Dispose of the document object to free resources
doc.Dispose();
```

---

# Word Document Header Management
## Add header only to the first page of a Word document
```javascript
// Load the source file
let doc1 = wasmModule.Document.Create();
doc1.LoadFromFile(inputFileName2);

// Get the header from the first section
let header = doc1.Sections.get(0).HeadersFooters.Header;

// Load the destination file
let doc2 = wasmModule.Document.Create();
doc2.LoadFromFile(inputFileName);

// Get the first page header of the destination document
let firstPageHeader = doc2.Sections.get(0).HeadersFooters.FirstPageHeader;

// Specify that the current section has a different header/footer for the first page
for (let i = 0; i < doc2.Sections.Count; i++) {
    let section = doc2.Sections.get_Item(i);
    section.PageSetup.DifferentFirstPageHeaderFooter = true;
}

// Removes all child objects in firstPageHeader
firstPageHeader.Paragraphs.Clear();

// Add all child objects of the header to firstPageHeader
for (let i = 0; i < header.ChildObjects.Count; i++) {
    let obj = header.ChildObjects.get(i);
    firstPageHeader.ChildObjects.Add(obj.Clone());
}
```

---

# Word Header and Footer Height Adjustment
## Adjust the height of headers and footers in a Word document
```javascript
//Load the document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Get the first section
let section = doc.Sections.get(0);

//Adjust the height of headers in the section
section.PageSetup.HeaderDistance = 100;

//Adjust the height of footers in the section
section.PageSetup.FooterDistance = 100;

//Save and launch document
doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
```

---

# Spire.Doc JavaScript Header and Footer
## Copy header between Word documents
```javascript
//Load the source file
let doc1 = wasmModule.Document.Create();
doc1.LoadFromFile(inputFileName);

//Get the header section from the source document
let header = doc1.Sections.get(0).HeadersFooters.Header;

//Load the destination file
let doc2 = wasmModule.Document.Create();
doc2.LoadFromFile(inputFileName_1);

//Copy each object in the header of source file to destination file
for (let i = 0; i < doc2.Sections.Count; i++) {
    let section = doc2.Sections.get_Item(i);
    for (let j = 0; j < header.ChildObjects.Count; j++) {
        let obj = header.ChildObjects.get(j);
        section.HeadersFooters.Header.ChildObjects.Add(obj.Clone());
    }
}
```

---

# Word Document Different First Page Header and Footer
## Set up different headers and footers for the first page of a Word document
```javascript
// Load the document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

// Get the section and set the property true
let section = doc.Sections.get_Item(0);
section.PageSetup.DifferentFirstPageHeaderFooter = true;

// Set the first page header. Here we append a picture in the header
let paragraph1 = section.HeadersFooters.FirstPageHeader.AddParagraph();
paragraph1.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;
let headerimage = paragraph1.AppendPicture({imgFile: inputImgFileName});

// Set the first page footer
let paragraph2 = section.HeadersFooters.FirstPageFooter.AddParagraph();
paragraph2.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
let FF = paragraph2.AppendText("First Page Footer");
FF.CharacterFormat.FontSize = 10;

// Set the other header & footer. If you only need the first page header & footer, don't set this
let paragraph3 = section.HeadersFooters.Header.AddParagraph();
paragraph3.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
let NH = paragraph3.AppendText("Spire.Doc for .NET");
NH.CharacterFormat.FontSize = 10;

let paragraph4 = section.HeadersFooters.Footer.AddParagraph();
paragraph4.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
let NF = paragraph4.AppendText("E-iceblue");
NF.CharacterFormat.FontSize = 10;
```

---

# Word Document Header and Footer
## Insert header and footer with images and page numbers
```javascript
function InsertHeaderAndFooter(section, inputImgFileName, inputImgFileName_1) {
  let header = section.HeadersFooters.Header;
  let footer = section.HeadersFooters.Footer;

  // Insert picture and text to header
  let headerParagraph = header.AddParagraph();

  let headerPicture = headerParagraph.AppendPicture({imgFile: inputImgFileName});
  // Header text
  let text = headerParagraph.AppendText("Demo of Spire.Doc");
  text.CharacterFormat.FontName = "Arial";
  text.CharacterFormat.FontSize = 10;
  text.CharacterFormat.Italic = true;
  headerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

  // Border
  headerParagraph.Format.Borders.Bottom.BorderType = wasmModule.BorderStyle.Single;
  headerParagraph.Format.Borders.Bottom.Space = 0.05;

  // Header picture layout - text wrapping
  headerPicture.TextWrappingStyle = wasmModule.TextWrappingStyle.Behind;

  // Header picture layout - position
  headerPicture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
  headerPicture.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
  headerPicture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
  headerPicture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Top;

  // Insert picture to footer
  let footerParagraph = footer.AddParagraph();

  let footerPicture = footerParagraph.AppendPicture({imgFile: inputImgFileName_1});

  // Footer picture layout
  footerPicture.TextWrappingStyle = wasmModule.TextWrappingStyle.Behind;
  footerPicture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
  footerPicture.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
  footerPicture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
  footerPicture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Bottom;

  // Insert page number
  footerParagraph.AppendField("page number", wasmModule.FieldType.FieldPage);
  footerParagraph.AppendText(" of ");
  footerParagraph.AppendField("number of pages", wasmModule.FieldType.FieldNumPages);
  footerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

  // Border
  footerParagraph.Format.Borders.Top.BorderType = wasmModule.BorderStyle.Single;
  footerParagraph.Format.Borders.Top.Space = 0.05;
}
```

---

# Spire.Doc JavaScript Header and Footer
## Add images to header and footer in a Word document
```javascript
// Get the header of the first page
let header = doc.Sections.get(0).HeadersFooters.Header;

// Add a paragraph for the header
let paragraph = header.AddParagraph();

// Set the format of the paragraph
paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

// Append a picture in the paragraph
let headerimage = paragraph.AppendPicture({imgFile: imageFileName});
headerimage.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Bottom;

// Get the footer of the first section
let footer = doc.Sections.get_Item(0).HeadersFooters.Footer;

// Add a paragraph for the footer
let paragraph2 = footer.AddParagraph();

// Set the format of the paragraph
paragraph2.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;

// Append a picture in the paragraph
let footerimage = paragraph2.AppendPicture({imgFile: logoFileName});

// Append text in the paragraph
let TR = paragraph2.AppendText("Copyright © 2013 e-iceblue. All Rights Reserved.");
TR.CharacterFormat.FontName = "Arial";
TR.CharacterFormat.FontSize = 10;
TR.CharacterFormat.TextColor = wasmModule.Color.get_Black();
```

---

# spire.doc javascript header
## lock header in word document
```javascript
//Load the document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Get the first section
let section = doc.Sections.get(0);

//Protect the document and set the ProtectionType as AllowOnlyFormFields
doc.Protect({type: wasmModule.ProtectionType.AllowOnlyFormFields, password: "123"});

//Set the ProtectForm as false to unprotect the section
section.ProtectForm = false;

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
```

---

# Spire.Doc JavaScript Header and Footer
## Add different headers and footers for odd and even pages
```javascript
//Load the document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Get the section
let section = doc.Sections.get_Item(0);

//Set the DifferentOddAndEvenPagesHeaderFooter property to true
section.PageSetup.DifferentOddAndEvenPagesHeaderFooter = true;

//Add odd header
let P3 = section.HeadersFooters.OddHeader.AddParagraph();
let OH = P3.AppendText("Odd Header");
P3.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
OH.CharacterFormat.FontName = "Arial";
OH.CharacterFormat.FontSize = 10;

//Add even header
let P4 = section.HeadersFooters.EvenHeader.AddParagraph();
let EH = P4.AppendText("Even Header from E-iceblue Using Spire.Doc");
P4.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
EH.CharacterFormat.FontName = "Arial";
EH.CharacterFormat.FontSize = 10;

//Add odd footer
let P2 = section.HeadersFooters.OddFooter.AddParagraph();
let OF = P2.AppendText("Odd Footer");
P2.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
OF.CharacterFormat.FontName = "Arial";
OF.CharacterFormat.FontSize = 10;

//Add even footer
let P1 = section.HeadersFooters.EvenFooter.AddParagraph();
let EF = P1.AppendText("Even Footer from E-iceblue Using Spire.Doc");
EF.CharacterFormat.FontName = "Arial";
EF.CharacterFormat.FontSize = 10;
P1.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
```

---

# Word Document Page Border Configuration
## Configure page border to surround header and footer in a Word document
```javascript
//Create a new document
let doc = wasmModule.Document.Create();
let section = doc.AddSection();

//Add a sample page border to the document
section.PageSetup.Borders.BorderType = wasmModule.BorderStyle.Wave;
section.PageSetup.Borders.Color = wasmModule.Color.get_Green();
section.PageSetup.Borders.Left.Space = 20;
section.PageSetup.Borders.Right.Space = 20;

//Add a header and set its format
let paragraph1 = section.HeadersFooters.Header.AddParagraph();
paragraph1.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;
let headerText = paragraph1.AppendText("Header isn't included in page border");
headerText.CharacterFormat.FontName = "Calibri";
headerText.CharacterFormat.FontSize = 20;
headerText.CharacterFormat.Bold = true;

//Add a footer and set its format
let paragraph2 = section.HeadersFooters.Footer.AddParagraph();
paragraph2.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;
let footerText = paragraph2.AppendText("Footer is included in page border");
footerText.CharacterFormat.FontName = "Calibri";
footerText.CharacterFormat.FontSize = 20;
footerText.CharacterFormat.Bold = true;

//Set the header not included in the page border while the footer included
section.PageSetup.PageBorderIncludeHeader = false;
section.PageSetup.HeaderDistance = 40;
section.PageSetup.PageBorderIncludeFooter = true;
section.PageSetup.FooterDistance = 40;
```

---

# spire.doc javascript footer
## remove footer from word document
```javascript
//Get the first section
let section = doc.Sections.get_Item(0);

//Clear footer in the first page
let footer;
footer = section.HeadersFooters.get_Item({hfType: wasmModule.HeaderFooterType.FooterFirstPage});
if (footer != null)
    footer.ChildObjects.Clear();
//Clear footer in the odd page
footer = section.HeadersFooters.get_Item({hfType: wasmModule.HeaderFooterType.FooterOdd});
if (footer != null)
    footer.ChildObjects.Clear();
//Clear footer in the even page
footer = section.HeadersFooters.get_Item({hfType: wasmModule.HeaderFooterType.FooterEven});
if (footer != null)
    footer.ChildObjects.Clear();
```

---

# Remove Headers from Word Document
## This code demonstrates how to remove headers from a Word document using Spire.Doc for JavaScript
```javascript
// Get the first section of the document
let section = doc.Sections.get_Item(0);

// Clear header in the first page
let header = section.HeadersFooters.get_Item({hfType: wasmModule.HeaderFooterType.HeaderFirstPage});
if (header != null)
    header.ChildObjects.Clear();

// Clear header in the odd page
header = section.HeadersFooters.get_Item({hfType: wasmModule.HeaderFooterType.HeaderOdd});
if (header != null)
    header.ChildObjects.Clear();

// Clear header in the even page
header = section.HeadersFooters.get_Item({hfType: wasmModule.HeaderFooterType.HeaderEven});
if (header != null)
    header.ChildObjects.Clear();
```

---

# spire.doc javascript table
## add alternative text to table in word document
```javascript
//Load the document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Get the first section
let section = doc.Sections.get(0);

//Get the first table in the section
let table = section.Tables.get_Item(0);

//Add alternative text
//Add title
table.Title = "Table 1";
//Add description
table.TableDescription = "Description Text";

// Save the document to the specified path
doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
```

---

# Spire.Doc JavaScript Table Row Operations
## Add or delete rows in a Word document table
```javascript
// Create a document
let document = wasmModule.Document.Create();
// Load file
document.LoadFromFile(inputFileName);
let section = document.Sections.get_Item(0);
let table = section.Tables.get_Item(0);

// Delete the seventh row
table.Rows.RemoveAt(7);

// Add a row and insert it into specific position
let row = wasmModule.TableRow.Create(document);
for (let i = 0; i < table.Rows.get(0).Cells.Count; i++) {
    let tc = row.AddCell();
    let paragraph = tc.AddParagraph();
    paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
    paragraph.AppendText("Added");
}
table.Rows.Insert(2, row);
// Add a row at the end of table
table.AddRow();
```

---

# spire.doc javascript table manipulation
## add or remove columns from a word table
```javascript
// Access the first section
let section = doc.Sections.get_Item(0);

// Access the first table
let table = section.Tables.get_Item(0);

// Add a blank column
let columnIndex1 = 0;
AddColumn(table, columnIndex1);

// Remove a column
let columnIndex2 = 2;
RemoveColumn(table, columnIndex2);

function AddColumn(table, columnIndex) {
    for (let r = 0; r < table.Rows.Count; r++) {
        let addCell = wasmModule.TableCell.Create(table.Document);
        table.Rows.get(r).Cells.Insert(columnIndex, addCell);
    }
}

function RemoveColumn(table, columnIndex) {
    for (let r = 0; r < table.Rows.Count; r++) {
        table.Rows.get(r).Cells.RemoveAt(columnIndex);
    }
}
```

---

# spire.doc javascript table
## add picture to table cell
```javascript
//Get the first table from the first section of the document
let table1 = doc.Sections.get_Item(0).Tables.get_Item(0);

//Add a picture to the specified table cell and set picture size
let picture = table1.Rows.get(1).Cells.get(2).Paragraphs.get_Item(0).AppendPicture({imgFile: inputImgFileName});

picture.Width = 100;
picture.Height = 100;
```

---

# Spire.Doc JavaScript Table
## Allow table rows to break across pages in Word document
```javascript
let section = document.Sections.get_Item(0);
let table = section.Tables.get_Item(0);

for (let i = 0; i < table.Rows.Count; i++) {
    let row = table.Rows.get_Item(i);
    //Allow break across pages
    row.RowFormat.IsBreakAcrossPages = true;
}
```

---

# Spire.Doc JavaScript Table AutoFit
## Automatically fit table columns to content in Word document
```javascript
//Create a document
let document = wasmModule.Document.Create();
//Load file
document.LoadFromFile(inputFileName);

let section = document.Sections.get_Item(0);
let table = section.Tables.get_Item(0);
//Automatically fit the table to the cell content
table.AutoFit(wasmModule.AutoFitBehaviorType.AutoFitToContents);
```

---

# spire.doc javascript table
## set table auto-fit to fixed column widths
```javascript
// Get the table from the document
let section = document.Sections.get_Item(0);
let table = section.Tables.get_Item(0);

// The table is set to a fixed size
table.AutoFit(wasmModule.AutoFitBehaviorType.FixedColumnWidths);
```

---

# spire.doc javascript table auto-fit
## automatically fit table to window width
```javascript
//Create a document
let document = wasmModule.Document.Create();
//Load file
document.LoadFromFile(inputFileName);
let section = document.Sections.get_Item(0);
let table = section.Tables.get_Item(0);
//Automatically fit the table to the active window width
table.AutoFit(wasmModule.AutoFitBehaviorType.AutoFitToWindow);
```

---

# Word Document Cell Merge Status Check
## Check if cells in a Word table are merged
```javascript
// Get the first section
let section = doc.Sections.get_Item(0);

// Get the first table in the section
let table = section.Tables.get_Item(0);

let stringBuidler = [];
for (let i = 0; i < table.Rows.Count; i++) {
    let tableRow = table.Rows.get(i);
    for (let j = 0; j < tableRow.Cells.Count; j++) {
        let tableCell = tableRow.Cells.get(j);
        let verticalMerge = tableCell.CellFormat.VerticalMerge;
        let horizontalMerge = tableCell.GridSpan;
        if (verticalMerge === wasmModule.CellMerge.None && horizontalMerge === 1) {
            stringBuidler.push("Row " + i + ", cell " + j + ": ");
            stringBuidler.push("This cell isn't merged.\n");
        } else {
            stringBuidler.push("Row " + i + ", cell " + j + ": ");
            stringBuidler.push("This cell is merged.\n");
        }
    }
    stringBuidler.push("\n");
}
```

---

# spire.doc javascript table
## clone a row in a word table
```javascript
//Get the first section
let se = doc.Sections.get_Item(0);

//Get the first row of the first table
let firstRow = se.Tables.get_Item(0).Rows.get_Item(0);

//Copy the first row to clone_FirstRow via TableRow.clone()
let clone_FirstRow = firstRow.Clone();

se.Tables.get_Item(0).Rows.Add(clone_FirstRow);
```

---

# spire.doc javascript table
## clone table in word document
```javascript
//Load the document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Get the first section
let se = doc.Sections.get_Item(0);

//Get the first table
let original_Table = se.Tables.get_Item(0);

//Copy the existing table to copied_Table via Table.clone()
let copied_Table = original_Table.Clone();
let st = ["Spire.Presentation for JS", "A professional " +
"PowerPoint® compatible library that enables developers to create, read, " +
"write, modify, convert and Print PowerPoint documents on any JS framework."];
//Get the last row of table
let lastRow = copied_Table.Rows.get(copied_Table.Rows.Count - 1);
//Change last row data
for (let i = 0; i < lastRow.Cells.Count - 1; i++) {
    lastRow.Cells.get(i).Paragraphs.get_Item(0).Text = st[i];
}
//Add copied_Table in section
se.Tables.Add(copied_Table);
```

---

# Word Document Table Manipulation
## Combine and split tables in a Word document
```javascript
// Get the first section
let section = doc.Sections.get_Item(0);

// Get the first and second table
let table1 = section.Tables.get_Item(0);
let table2 = section.Tables.get_Item(1);

// Add the rows of table2 to table1
for (let i = 0; i < table2.Rows.Count; i++) {
    table1.Rows.Add(table2.Rows.get(i).Clone());
}

// Remove the table2
section.Tables.Remove(table2);

// Split table function
function SplitTable(inputFileName) {
    // Get the first section
    let section = doc.Sections.get(0);

    // Get the first table
    let table = section.Tables.get_Item(0);

    // We will split the table at the third row
    let splitIndex = 2;

    // Create a new table for the split table
    let newTable = wasmModule.Table.Create(section.Document, false);

    // Add rows to the new table
    for (let i = splitIndex; i < table.Rows.Count; i++) {
        newTable.Rows.Add(table.Rows.get(i).Clone());
    }

    // Remove rows from original table
    for (let i = table.Rows.Count - 1; i >= splitIndex; i--) {
        table.Rows.RemoveAt(i);
    }

    // Add the new table in section
    section.Tables.Add(newTable);
}
```

---

# Spire.Doc JavaScript Nested Table
## Create a nested table in a Word document
```javascript
//Create a new document
let doc = wasmModule.Document.Create();
let section = doc.AddSection();

//Add a table
let table = section.AddTable({showBorder: true});
table.ResetCells(2, 2);

//Set column width
table.Rows.get(0).Cells.get(0).SetCellWidth(70, wasmModule.CellWidthType.Point);
table.Rows.get(1).Cells.get(0).SetCellWidth(70, wasmModule.CellWidthType.Point);
table.AutoFit(wasmModule.AutoFitBehaviorType.AutoFitToWindow);

//Insert content to cells
table.Rows.get(0).Cells.get(0).AddParagraph().AppendText("Spire.Doc for .NET");
let text = "Spire.Doc for .NET is a professional Word" +
    ".NET library specifically designed for developers to create," +
    "read, write, convert and print Word document files from any .NET" +
    "platform with fast and high quality performance.";
table.Rows.get(0).Cells.get(1).AddParagraph().AppendText(text);

//Add a nested table to cell(first row, second column)
let nestedTable = table.Rows.get(0).Cells.get(1).AddTable({showBorder: true});
nestedTable.ResetCells(4, 3);
nestedTable.AutoFit(wasmModule.AutoFitBehaviorType.AutoFitToContents);

//Add content to nested cells
nestedTable.Rows.get(0).Cells.get(0).AddParagraph().AppendText("NO.");
nestedTable.Rows.get(0).Cells.get(1).AddParagraph().AppendText("Item");
nestedTable.Rows.get(0).Cells.get(2).AddParagraph().AppendText("Price");

nestedTable.Rows.get(1).Cells.get(0).AddParagraph().AppendText("1");
nestedTable.Rows.get(1).Cells.get(1).AddParagraph().AppendText("Pro Edition");
nestedTable.Rows.get(1).Cells.get(2).AddParagraph().AppendText("$799");

nestedTable.Rows.get(2).Cells.get(0).AddParagraph().AppendText("2");
nestedTable.Rows.get(2).Cells.get(1).AddParagraph().AppendText("Standard Edition");
nestedTable.Rows.get(2).Cells.get(2).AddParagraph().AppendText("$599");

nestedTable.Rows.get(3).Cells.get(0).AddParagraph().AppendText("3");
nestedTable.Rows.get(3).Cells.get(1).AddParagraph().AppendText("Free Edition");
nestedTable.Rows.get(3).Cells.get(2).AddParagraph().AppendText("$0");
```

---

# Spire.Doc JavaScript Table
## Create Table in Word Document
```javascript
function addTable(section) {
    let table = section.AddTable({showBorder: true});
    table.ResetCells(data.length + 1, header.length);

    // ***************** First Row *************************
    let row = table.Rows.get_Item(0);
    row.IsHeader = true;
    row.Height = 20;    //unit: point, 1point = 0.3528 mm
    row.HeightType = wasmModule.TableRowHeightType.Exactly;
    row.RowFormat.BackColor = wasmModule.Color.get_Gray();
    for (let i = 0; i < header.length; i++) {
        row.Cells.get(i).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
        let p = row.Cells.get(i).AddParagraph();
        p.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        let txtRange = p.AppendText(header[i]);
        txtRange.CharacterFormat.Bold = true;
    }

    for (let r = 0; r < data.length; r++) {
        let dataRow = table.Rows.get(r + 1);
        dataRow.Height = 20;
        dataRow.HeightType = wasmModule.TableRowHeightType.Exactly;
        dataRow.RowFormat.BackColor = wasmModule.Color.Empty();
        for (let c = 0; c < data[r].length; c++) {
            dataRow.Cells.get(c).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
            dataRow.Cells.get(c).AddParagraph().AppendText(data[r][c]);
        }
    }

    for (let j = 1; j < table.Rows.Count; j++) {
        if (j % 2 == 0) {
            let row2 = table.Rows.get_Item(j);
            for (let f = 0; f < row2.Cells.Count; f++) {
                row2.Cells.get(f).CellFormat.BackColor = wasmModule.Color.get_LightBlue();
            }
        }
    }
}
```

---

# spire.doc javascript table
## create table directly in word document
```javascript
//Create a Word document
let doc = wasmModule.Document.Create();

//Add a section
let section = doc.AddSection();

//Create a table
let table = wasmModule.Table.Create(doc, false);
//Set the width of table
table.PreferredWidth = wasmModule.PreferredWidth.Create(wasmModule.WidthType.Percentage, 100);
//Set the border of table
table.TableFormat.Borders.BorderType = wasmModule.BorderStyle.Single;

//Create a table row
let row = wasmModule.TableRow.Create(doc, false);
row.Height = 50.0;
table.Rows.Add(row);

//Create a table cell
let cell1 = wasmModule.TableCell.Create(doc);
//Add a paragraph
let para1 = cell1.AddParagraph();
//Append text in the paragraph
para1.AppendText("Row 1, Cell 1");
//Set the horizontal alignment of paragraph
para1.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
//Set the background color of cell
cell1.CellFormat.BackColor = wasmModule.Color.get_CadetBlue();
//Set the vertical alignment of paragraph
cell1.CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
row.Cells.Add(cell1);

//Create a table cell
let cell2 = wasmModule.TableCell.Create(doc);
let para2 = cell2.AddParagraph();
para2.AppendText("Row 1, Cell 2");
para2.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
cell2.CellFormat.BackColor = wasmModule.Color.get_CadetBlue();
cell2.CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
row.Cells.Add(cell2);

//Add the table in the section
section.Tables.Add(table);
```

---

# spire.doc javascript table
## create table from html in word document
```javascript
//HTML string
let HTML = "<table border='2px'>" +
    "<tr>" +
    "<td>Row 1, Cell 1</td>" +
    "<td>Row 1, Cell 2</td>" +
    "</tr>" +
    "<tr>" +
    "<td>Row 2, Cell 2</td>" +
    "<td>Row 2, Cell 2</td>" +
    "</tr>" +
    "</table>";

//Create a Word document
let document = wasmModule.Document.Create();

//Add a section
let section = document.AddSection();

//Add a paragraph and append html string
section.AddParagraph().AppendHTML(HTML);

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
```

---

# Spire.Doc JavaScript Vertical Table
## Create a vertical table in a Word document
```javascript
// Create Word document
let document = wasmModule.Document.Create();

// Add a new section
let section = document.AddSection();

// Add a table with rows and columns and set the text for the table
let table = section.AddTable();
table.ResetCells(1, 1);
let cell = table.Rows.get(0).Cells.get(0);
table.Rows.get(0).Height = 150;
cell.AddParagraph().AppendText("Draft copy in vertical style");

// Set the TextDirection for the table to RightToLeftRotated
cell.CellFormat.TextDirection = wasmModule.TextDirection.RightToLeftRotated;

// Set the table format
table.TableFormat.WrapTextAround = true;
table.TableFormat.Positioning.VertRelationTo = wasmModule.VerticalRelation.Page;
table.TableFormat.Positioning.HorizRelationTo = wasmModule.HorizontalRelation.Page;
table.TableFormat.Positioning.HorizPosition = section.PageSetup.PageSize.Width - table.Width;
table.TableFormat.Positioning.VertPosition = 200;
```

---

# Spire.Doc JavaScript Table Borders
## Set different borders on tables and cells in Word documents
```javascript
function setTableBorders(table) {
    table.TableFormat.Borders.BorderType = wasmModule.BorderStyle.Single;
    table.TableFormat.Borders.LineWidth = 3.0;
    table.TableFormat.Borders.Color = wasmModule.Color.get_Red();
}

function setCellBorders(tableCell) {
    tableCell.CellFormat.Borders.BorderType = wasmModule.BorderStyle.DotDash;
    tableCell.CellFormat.Borders.LineWidth = 1.0;
    tableCell.CellFormat.Borders.Color = wasmModule.Color.get_Green();
}

// Get table from document
let table = document.Sections.get_Item(0).Tables.get_Item(0);

// Set borders of table
setTableBorders(table);

// Set borders of cell
setCellBorders(table.Rows.get(2).Cells.get(0));
```

---

# Word Document Table Merged Cells Formatting
## Format merged cells in a Word document table with custom styles and alignment
```javascript
// Create word document
let document = wasmModule.Document.Create();

// Add a new section
let section = document.AddSection();

// Add a new table
let table = AddTable(section);

// Create a new style
let style = wasmModule.ParagraphStyle.Create(document);
style.Name = "Style";
style.CharacterFormat.TextColor = wasmModule.Color.get_DeepSkyBlue();
style.CharacterFormat.Italic = true;
style.CharacterFormat.Bold = true;
style.CharacterFormat.FontSize = 13;
document.Styles.Add(style);

// Merge cell horizontally
table.ApplyHorizontalMerge(0, 0, 1);
// Apply style
table.Rows.get(0).Cells.get(0).Paragraphs.get_Item(0).ApplyStyle(style.Name);
// Set vertical and horizontal alignment
table.Rows.get(0).Cells.get(0).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
table.Rows.get(0).Cells.get(0).Paragraphs.get_Item(0).Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

// Merge cell vertically
table.ApplyVerticalMerge(0, 1, 3);
// Apply style
table.Rows.get(1).Cells.get(0).Paragraphs.get_Item(0).ApplyStyle(style.Name);
// Set vertical and horizontal alignment
table.Rows.get(1).Cells.get(0).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
table.Rows.get(1).Cells.get(0).Paragraphs.get_Item(0).Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;
// Set column width
table.Rows.get(1).Cells.get(0).SetCellWidth(20, wasmModule.CellWidthType.Percentage);

function AddTable(section) {
    let table = section.AddTable({showBorder: true});
    table.ResetCells(4, 3);
    // Table data
    let dt = [["Product", "", "Price"],
        ["Spire.Doc", "Pro Edition", "$799"],
        ["", "Standard Edition", "$599"],
        ["", "Free Edition", "$0"]];

    for (let r = 0; r < dt.length; r++) {
        let dataRow = table.Rows.get(r);
        dataRow.Height = 20;
        dataRow.HeightType = wasmModule.TableRowHeightType.Exactly;
        dataRow.RowFormat.BackColor = wasmModule.Color.Empty;
        for (let c = 0; c < dataRow.Cells.Count; c++) {
            if(dt[r][c] !==""){
                let range = dataRow.Cells.get(c).AddParagraph().AppendText(dt[r][c]);
                range.CharacterFormat.FontName = "Arial";
            }
        }
    }
    return table;
}
```

---

# Spire.Doc JavaScript Table Diagonal Border
## Get diagonal border properties of table cell in Word document
```javascript
//Load Word from disk
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Get the first section
let section = doc.Sections.get(0);

//Get the first table in the section
let table = section.Tables.get_Item(0);

//Get the setting of the diagonal border of table cell
let bs_UP = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalUp.BorderType;
let color_UP = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalUp.Color;
let width_UP = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalUp.LineWidth;
let bs_Down = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalDown.BorderType;
let color_Down = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalDown.Color;
let width_Down = table.Rows.get(0).Cells.get(0).CellFormat.Borders.DiagonalDown.LineWidth;
```

---

# spire.doc javascript table
## get row and cell index of table
```javascript
//Load Word from disk
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Get the first section
let section = doc.Sections.get_Item(0);

//Get the first table in the section
let table = section.Tables.get_Item(0);

//Get table collections
let collections = section.Tables;

//Get the table index
let tableIndex = collections.IndexOf(table);

//Get the index of the last table row
let row = table.LastRow;
let rowIndex = row.GetRowIndex();

//Get the index of the last table cell
let cell = row.LastChild;
let cellIndex = cell.GetCellIndex();
```

---

# spire.doc javascript table
## get table position in word document
```javascript
//Get the first section
let section = document.Sections.get_Item(0);
//Get the first table
let table = section.Tables.get_Item(0);

//Verify whether the table uses "Around" text wrapping or not.
if (table.TableFormat.WrapTextAround) {
    let positon = table.TableFormat.Positioning;
    
    // Get horizontal position information
    let horizPosition = positon.HorizPosition;
    let horizPositionAbs = positon.HorizPositionAbs;
    let horizRelationTo = positon.HorizRelationTo;
    
    // Get vertical position information
    let vertPosition = positon.VertPosition;
    let vertPositionAbs = positon.VertPositionAbs;
    let vertRelationTo = positon.VertRelationTo;
    
    // Get distance from surrounding text
    let distanceFromTop = positon.DistanceFromTop;
    let distanceFromLeft = positon.DistanceFromLeft;
    let distanceFromBottom = positon.DistanceFromBottom;
    let distanceFromRight = positon.DistanceFromRight;
}
```

---

# Word Document Table Cell Operations
## Merge and split table cells in a Word document
```javascript
// Create a document and load file from disk
let document = wasmModule.Document.Create();
document.LoadFromFile(inputFileName);
let section = document.Sections.get_Item(0);
let table = section.Tables.get_Item(0);

// The method shows how to merge cell horizontally
table.ApplyHorizontalMerge(6, 2, 3);

// The method shows how to merge cell vertically
table.ApplyVerticalMerge(2, 4, 5);

// The method shows how to split the cell
table.Rows.get(8).Cells.get(3).SplitCell(2, 2);
```

---

# Spire.Doc JavaScript Table Format Modification
## Modify table format including row format and cell format
```javascript
// Get the first section
let section = document.Sections.get(0);

// Get tables
let tb1 = section.Tables.get_Item(0);
let tb2 = section.Tables.get_Item(1);
let tb3 = section.Tables.get_Item(2);

MoidyTableFormat(tb1);
ModifyRowFormat(tb2);
ModifyCellFormat(tb3);

function MoidyTableFormat(table) {
    // Set table width
    table.PreferredWidth = wasmModule.PreferredWidth.Create(wasmModule.WidthType.Twip, 6000);

    // Apply style for table
    table.ApplyStyle(wasmModule.DefaultTableStyle.ColorfulGridAccent3);

    // Set table padding
    table.TableFormat.Paddings.All = 5;

    // Set table title and description
    table.Title = "Spire.Doc for .NET";
    table.TableDescription = "Spire.Doc for .NET is a professional Word .NET library";
}

function ModifyRowFormat(table) {
    // Set cell spacing
    table.Rows.get(0).RowFormat.CellSpacing = 2;

    // Set row height
    table.Rows.get(1).HeightType = wasmModule.TableRowHeightType.Exactly;
    table.Rows.get(1).Height = 20;

    // Set background color
    table.Rows.get(2).RowFormat.BackColor = wasmModule.Color.get_DarkSeaGreen();
}

function ModifyCellFormat(table) {
    // Set alignment
    table.Rows.get(0).Cells.get(0).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
    table.Rows.get(0).Cells.get(0).Paragraphs.get_Item(0).Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

    // Set background color
    table.Rows.get(1).Cells.get(0).CellFormat.BackColor = wasmModule.Color.get_DarkSeaGreen();

    // Set cell border
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.BorderType = wasmModule.BorderStyle.Single;
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.LineWidth = 1;
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.Left.Color = wasmModule.Color.get_Red();
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.Right.Color = wasmModule.Color.get_Red();
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.Top.Color = wasmModule.Color.get_Red();
    table.Rows.get(2).Cells.get(0).CellFormat.Borders.Bottom.Color = wasmModule.Color.get_Red();

    // Set text direction
    table.Rows.get(3).Cells.get(0).CellFormat.TextDirection = wasmModule.TextDirection.RightToLeft;
}
```

---

# spire.doc javascript table
## prevent page break in table
```javascript
//Get the table from Word document.
let table = document.Sections.get(0).Tables.get_Item(0);

//Change the paragraph setting to keep them together.
for (let i = 0; i < table.Rows.Count; i++) {
    let row = table.Rows.get_Item(i);
    for (let j = 0; j < row.Cells.Count; j++) {
        let cell = row.Cells.get(j);
        for (let k = 0; k < cell.Paragraphs.Count; k++) {
            let p = cell.Paragraphs.get_Item(k);
            p.Format.KeepFollow = true;
        }
    }
}
```

---

# Remove Table from Word Document
## This code demonstrates how to remove a table from a Word document using JavaScript
```javascript
// Load the document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

// Remove the first Table
doc.Sections.get(0).Tables.RemoveAt(0);

// Save the document
doc.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013});

// Clean up resources
doc.Close();
doc.Dispose();
```

---

# Spire.Doc JavaScript Table Header
## Configure table rows to repeat on each page
```javascript
//Create a table
let table = section.AddTable({showBorder: true});

//Add a header row that will repeat on each page
let row = table.AddRow();
//Set the row as a table header to repeat on each page
row.IsHeader = true;
//Set the backcolor of row
row.RowFormat.BackColor = wasmModule.Color.get_LightGray();
//Add a new cell for row
let cell = row.AddCell();
cell.SetCellWidth(100, wasmModule.CellWidthType.Percentage);
//Add a paragraph for cell to put some data
let parapraph = cell.AddParagraph();
//Add text
parapraph.AppendText("Row Header 1");
//Set paragraph horizontal center alignment
parapraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

//Add another header row that will repeat on each page
row = table.AddRow({isCopyFormat: false, columnsNum: 1});
row.IsHeader = true;
row.RowFormat.BackColor = wasmModule.Color.get_Ivory();
//Set row height
row.Height = 30;
cell = row.Cells.get(0);
cell.SetCellWidth(100, wasmModule.CellWidthType.Percentage);
//Set cell vertical middle alignment
cell.CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
//Add a paragraph for cell to put some data
parapraph = cell.AddParagraph();
//Add text
parapraph.AppendText("Row Header 2");
parapraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
```

---

# spire.doc javascript table text replacement
## replace text in word table using regex and string matching
```javascript
//Get the first section
let section = doc.Sections.get_Item(0);

//Get the first table in the section
let table = section.Tables.get_Item(0);

//Define a regular expression to match the {} with its content
let regex = wasmModule.Regex.Create("{[^\\}]+\\}", wasmModule.RegexOptions.None);

//Replace the text of table with regex
table.Replace({pattern: regex, replace: "E-iceblue"});

//Replace old text with new text in table
table.Replace({given:"Beijing", replace: "Component", caseSensitive: false, wholeWord: true});
```

---

# Setting Table Column Width
## This code demonstrates how to set the width of a column in a Word document table
```javascript
let section = document.Sections.get_Item(0);
let table = section.Tables.get_Item(0);

//Traverse the first column
for (let i = 0; i < table.Rows.Count; i++) {
    //Set the width and type of the cell
    table.Rows.get(i).Cells.get(0).SetCellWidth(200, wasmModule.CellWidthType.Point);
}
```

---

# spire.doc javascript table positioning
## set table position to outside in word document
```javascript
//Add a table of 4 rows and 2 columns
let table = header.AddTable();
table.ResetCells(4, 2);

//Set the position of the table to the right of the image
table.TableFormat.WrapTextAround = true;
table.TableFormat.Positioning.HorizPositionAbs = wasmModule.HorizontalPosition.Outside;
table.TableFormat.Positioning.VertRelationTo = wasmModule.VerticalRelation.Margin;
table.TableFormat.Positioning.VertPosition = 43;
```

---

# Spire.Doc JavaScript Table Styling
## Set table style and borders in a Word document
```javascript
let section = document.Sections.get(0);

//Get the first table
let table = section.Tables.get_Item(0);

//Apply the table style
table.ApplyStyle(wasmModule.DefaultTableStyle.ColorfulList);

//Set right border of table
table.TableFormat.Borders.Right.BorderType = wasmModule.BorderStyle.Hairline;
table.TableFormat.Borders.Right.LineWidth = 1.0;
table.TableFormat.Borders.Right.Color = wasmModule.Color.get_Red();

//Set top border of table
table.TableFormat.Borders.Top.BorderType = wasmModule.BorderStyle.Hairline;
table.TableFormat.Borders.Top.LineWidth = 1.0;
table.TableFormat.Borders.Top.Color = wasmModule.Color.get_Green();

//Set left border of table
table.TableFormat.Borders.Left.BorderType = wasmModule.BorderStyle.Hairline;
table.TableFormat.Borders.Left.LineWidth = 1.0;
table.TableFormat.Borders.Left.Color = wasmModule.Color.get_Yellow();

//Set bottom border is none
table.TableFormat.Borders.Bottom.BorderType = wasmModule.BorderStyle.DotDash;

//Set vertical and horizontal border
table.TableFormat.Borders.Vertical.BorderType = wasmModule.BorderStyle.Dot;
table.TableFormat.Borders.Horizontal.BorderType = wasmModule.BorderStyle.None;
table.TableFormat.Borders.Vertical.Color = wasmModule.Color.get_Orange();
```

---

# Spire.Doc JavaScript Table
## Set vertical alignment for table cells
```javascript
// Create a new Word document and add a new section
let doc = wasmModule.Document.Create();
let section = doc.AddSection();

// Add a table with 3 columns and 3 rows
let table = section.AddTable({showBorder: true});
table.ResetCells(3, 3);

// Merge rows
table.ApplyVerticalMerge(0, 0, 2);

// Set the vertical alignment for each cell, default is top
table.Rows.get(0).Cells.get(0).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
table.Rows.get(0).Cells.get(1).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Top;
table.Rows.get(0).Cells.get(2).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Top;
table.Rows.get(1).Cells.get(1).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
table.Rows.get(1).Cells.get(2).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
table.Rows.get(2).Cells.get(1).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Bottom;
table.Rows.get(2).Cells.get(2).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Bottom;

// Insert a picture into the table cell
let paraPic = table.Rows.get(0).Cells.get(0).AddParagraph();
let pic = paraPic.AppendPicture({imgFile: inputFileName});
```

---

# Spire.Doc JavaScript Image Hyperlink
## Create an image hyperlink in a Word document
```javascript
// Load Document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

let section = doc.Sections.get_Item(0);
// Add a paragraph
let paragraph = section.AddParagraph();
// Load an image to a DocPicture object

let picture = wasmModule.DocPicture.Create(doc);
// Add an image hyperlink to the paragraph
picture.LoadImage(inputFileName_1);

paragraph.AppendHyperlink({
    link: "https://www.e-iceblue.com/Introduce/word-for-net-introduce.html",
    picture: picture,
    type: wasmModule.HyperlinkType.WebLink
});
```

---

# spire.doc javascript hyperlinks
## find all hyperlinks in a word document
```javascript
//Load Document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Create a hyperlink list
let hyperlinks = [];
let hyperlinksText = [];
//Iterate through the items in the sections to find all hyperlinks
for (let i = 0; i < doc.Sections.Count; i++) {
    let section = doc.Sections.get(i);
    for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
        let sec = section.Body.ChildObjects.get(j);
        if (sec.DocumentObjectType == wasmModule.DocumentObjectType.Paragraph) {
            for (let k = 0; k < sec.ChildObjects.Count; k++) {
                let para = sec.ChildObjects.get(k);
                if (para.DocumentObjectType == wasmModule.DocumentObjectType.Field) {
                    let field = para;
                    if (field.Type == wasmModule.FieldType.FieldHyperlink) {
                        hyperlinks.push(field);
                        //Get the hyperlink text
                        hyperlinksText.push(field.FieldText + "\r\n");
                    }
                }
            }
        }
    }
}
```

---

# Spire.Doc JavaScript Hyperlink
## Insert hyperlinks in a Word document
```javascript
// Create a new Word document
let document = wasmModule.Document.Create();
let section = document.AddSection();

// Insert hyperlink
InsertHyperlink(section, inputFileName);

function InsertHyperlink(section, inputFileName) {
    let paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
    paragraph.AppendText("Spire.Doc for JS\r\n e-iceblue company Ltd. 2002-2010 All rights reserverd");
    paragraph.ApplyStyle({ builtinStyle: wasmModule.BuiltinStyle.Heading2 });

    paragraph = section.AddParagraph();
    paragraph.AppendText("Home page");
    paragraph.ApplyStyle({ builtinStyle: wasmModule.BuiltinStyle.Heading2 });
    paragraph = section.AddParagraph();
    paragraph.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", wasmModule.HyperlinkType.WebLink);

    paragraph = section.AddParagraph();
    paragraph.AppendText("Contact US");
    paragraph.ApplyStyle({ builtinStyle: wasmModule.BuiltinStyle.Heading2 });
    paragraph = section.AddParagraph();
    paragraph.AppendHyperlink("mailto:support@e-iceblue.com", "support@e-iceblue.com", wasmModule.HyperlinkType.EMailLink);

    paragraph = section.AddParagraph();
    paragraph.AppendText("Forum");
    paragraph.ApplyStyle({ builtinStyle: wasmModule.BuiltinStyle.Heading2 });
    paragraph = section.AddParagraph();
    paragraph.AppendHyperlink("www.e-iceblue.com/forum/", "www.e-iceblue.com/forum/", wasmModule.HyperlinkType.WebLink);

    paragraph = section.AddParagraph();
    paragraph.AppendText("Download Link");
    paragraph.ApplyStyle({ builtinStyle: wasmModule.BuiltinStyle.Heading2 });
    paragraph = section.AddParagraph();
    paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", "www.e-iceblue.com/Download/download-word-for-net-now.html", wasmModule.HyperlinkType.WebLink);

    paragraph = section.AddParagraph();
    paragraph.AppendText("Insert Link On Image");
    paragraph.ApplyStyle({ builtinStyle: wasmModule.BuiltinStyle.Heading2 });
    paragraph = section.AddParagraph();

    let picture = paragraph.AppendPicture({ imgFile: inputFileName });
    paragraph.AppendHyperlink("www.e-iceblue.com/Download/download-word-for-net-now.html", picture, wasmModule.HyperlinkType.WebLink);
}
```

---

# Spire.Doc JavaScript Hyperlink
## Modify hyperlink text in Word document
```javascript
//Load Document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);

//Find all hyperlinks in the Word document
let hyperlinks = [];
for (let i = 0; i < doc.Sections.Count; i++) {
    let section = doc.Sections.get(i);
    for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
        let sec = section.Body.ChildObjects.get(j);
        if (sec.DocumentObjectType == wasmModule.DocumentObjectType.Paragraph) {
            for (let k = 0; k < sec.ChildObjects.Count; k++) {
                let para = sec.ChildObjects.get(k);
                if (para.DocumentObjectType == wasmModule.DocumentObjectType.Field) {
                    let field = para;

                    if (field.Type == wasmModule.FieldType.FieldHyperlink) {
                        hyperlinks.push(field);
                    }
                }
            }
        }
    }
}

//Reset the property of hyperlinks[0].FieldText by using the index of the hyperlink
hyperlinks[0].FieldText = "Spire.Doc component";
```

---

# spire.doc javascript hyperlinks
## remove hyperlinks from word document
```javascript
//Get all hyperlinks
let hyperlinks = FindAllHyperlinks(doc);

//Flatten all hyperlinks
for (let i = hyperlinks.length - 1; i >= 0; i--) {
    FlattenHyperlinks(hyperlinks[i]);
}

function FindAllHyperlinks(document) {
    let hyperlinks = [];
    //Iterate through the items in the sections to find all hyperlinks
    for (let i = 0; i < document.Sections.Count; i++) {
        let section = document.Sections.get(i);
        for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
            let sec = section.Body.ChildObjects.get(j);
            if (sec.DocumentObjectType == wasmModule.DocumentObjectType.Paragraph) {
                for (let k = 0; k < sec.ChildObjects.Count; k++) {
                    let para = sec.ChildObjects.get(k);
                    if (para.DocumentObjectType == wasmModule.DocumentObjectType.Field) {
                        let field = para;

                        if (field.Type == wasmModule.FieldType.FieldHyperlink) {
                            hyperlinks.push(field);
                        }
                    }
                }
            }
        }
    }
    return hyperlinks;
}

function FlattenHyperlinks(field) {
    let ownerParaIndex = field.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.OwnerParagraph);
    let fieldIndex = field.OwnerParagraph.ChildObjects.IndexOf(field);
    let sepOwnerPara = field.Separator.OwnerParagraph;
    let sepOwnerParaIndex = field.Separator.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.Separator.OwnerParagraph);
    let sepIndex = field.Separator.OwnerParagraph.ChildObjects.IndexOf(field.Separator);
    let endIndex = field.End.OwnerParagraph.ChildObjects.IndexOf(field.End);
    let endOwnerParaIndex = field.End.OwnerParagraph.OwnerTextBody.ChildObjects.IndexOf(field.End.OwnerParagraph);

    FormatFieldResultText(field.Separator.OwnerParagraph.OwnerTextBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex);

    field.End.OwnerParagraph.ChildObjects.RemoveAt(endIndex);

    for (let i = sepOwnerParaIndex; i >= ownerParaIndex; i--) {
        if (i == sepOwnerParaIndex && i == ownerParaIndex) {
            for (let j = sepIndex; j >= fieldIndex; j--) {
                field.OwnerParagraph.ChildObjects.RemoveAt(j);

            }
        } else if (i == ownerParaIndex) {
            for (let j = field.OwnerParagraph.ChildObjects.Count - 1; j >= fieldIndex; j--) {
                field.OwnerParagraph.ChildObjects.RemoveAt(j);
            }

        } else if (i == sepOwnerParaIndex) {
            for (let j = sepIndex; j >= 0; j--) {
                sepOwnerPara.ChildObjects.RemoveAt(j);
            }
        } else {
            field.OwnerParagraph.OwnerTextBody.ChildObjects.RemoveAt(i);
        }
    }
}

function FormatFieldResultText(ownerBody, sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex) {
    for (let i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++) {
        let para = ownerBody.ChildObjects.get(i);
        if (i == sepOwnerParaIndex && i == endOwnerParaIndex) {
            for (let j = sepIndex + 1; j < endIndex; j++) {
                FormatText(para.ChildObjects.get(j));
            }

        } else if (i == sepOwnerParaIndex) {
            for (let j = sepIndex + 1; j < para.ChildObjects.Count; j++) {
                FormatText(para.ChildObjects.get(j));
            }
        } else if (i == endOwnerParaIndex) {
            for (let j = 0; j < endIndex; j++) {
                FormatText(para.ChildObjects.get(j));
            }
        } else {
            for (let j = 0; j < para.ChildObjects.Count; j++) {
                FormatText(para.ChildObjects.get(j));
            }
        }
    }
}

function FormatText(tr) {
    //Set the text color to black
    tr.CharacterFormat.TextColor = wasmModule.Color.get_Black();
    //Set the text underline style to none
    tr.CharacterFormat.UnderlineStyle = wasmModule.UnderlineStyle.None;
}
```

---

# spire.doc javascript hyperlink formatting
## change hyperlink color and remove hyperlink underline
```javascript
// Load Document
let doc = wasmModule.Document.Create();
doc.LoadFromFile(inputFileName);
let section = doc.Sections.get_Item(0);

// Add a paragraph and append a hyperlink to the paragraph
let para1 = section.AddParagraph();
para1.AppendText("Regular Link: ");
// Format the hyperlink with default color and underline style
let txtRange1 = para1.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", wasmModule.HyperlinkType.WebLink);
txtRange1.CharacterFormat.FontName = "Times New Roman";
txtRange1.CharacterFormat.FontSize = 12;
let blankPara1 = section.AddParagraph();

// Add a paragraph and append a hyperlink to the paragraph
let para2 = section.AddParagraph();
para2.AppendText("Change Color: ");
// Format the hyperlink with red color and underline style
let txtRange2 = para2.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", wasmModule.HyperlinkType.WebLink);
txtRange2.CharacterFormat.FontName = "Times New Roman";
txtRange2.CharacterFormat.FontSize = 12;
txtRange2.CharacterFormat.TextColor = wasmModule.Color.get_Red();
let blankPara2 = section.AddParagraph();

// Add a paragraph and append a hyperlink to the paragraph
let para3 = section.AddParagraph();
para3.AppendText("Remove Underline: ");
// Format the hyperlink with red color and no underline style
let txtRange3 = para3.AppendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", wasmModule.HyperlinkType.WebLink);
txtRange3.CharacterFormat.FontName = "Times New Roman";
txtRange3.CharacterFormat.FontSize = 12;
txtRange3.CharacterFormat.UnderlineStyle = wasmModule.UnderlineStyle.None;
```

---

# Word Document Decryption
## Remove password protection from a Word document
```javascript
// Load a document from the virtual file system
document.LoadFromFile({fileName: inputFileName, fileFormat: wasmModule.FileFormat.Docx, password: "E-iceblue"});

// Remove the encryption
document.RemoveEncryption();
```

---

# spire.doc javascript encryption
## encrypt Word document with password
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Encrypt document with password
document.Encrypt("E-iceblue");
```

---

# Word Document Section Locking
## Lock specified sections in a Word document while protecting the document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Add new sections
let s1 = document.AddSection();
let s2 = document.AddSection();

// Append some text to section 1 and section 2
s1.AddParagraph().AppendText("Spire.Doc demo, section 1");
s2.AddParagraph().AppendText("Spire.Doc demo, section 2");

// Protect the document with AllowOnlyFormFields protection type
document.Protect({type: wasmModule.ProtectionType.AllowOnlyFormFields, password: "123"});

// Unprotect section 2
s2.ProtectForm = false;
```

---

# spire.doc javascript security
## remove editable range from document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Find "PermissionStart" and "PermissionEnd" tags and remove them
for (let i = 0; i < document.Sections.Count; i++) {
    let section = document.Sections.get_Item(i);
    for (let j = 0; j < section.Body.Paragraphs.Count; j++) {
        let paragraph = section.Body.Paragraphs.get_Item(j);
        for (let i = 0; i < paragraph.ChildObjects.Count;) {
            let obj = paragraph.ChildObjects.get(i);
            if (obj instanceof wasmModule.PermissionStart || obj instanceof wasmModule.PermissionEnd) {
                paragraph.ChildObjects.Remove(obj);
            } else {
                i++;
            }
        }
    }
}
```

---

# Spire.Doc JavaScript Security
## Remove Read-Only Restriction from Word Document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document
document.LoadFromFile(inputFileName);

// Remove ReadOnly Restriction
document.Protect({type: wasmModule.ProtectionType.NoProtection});

// Save the document
document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
```

---

# Document Security - Editable Ranges
## Set editable ranges in a protected Word document
```javascript
// Protect whole document
document.Protect({type: wasmModule.ProtectionType.AllowOnlyReading, password: "password"});

// Create tags for permission start and end
let start = wasmModule.PermissionStart.Create(document, "testID");
let end = wasmModule.PermissionEnd.Create(document, "testID");

// Add the start and end tags to allow the first paragraph to be edited
document.Sections.get(0).Paragraphs.get_Item(0).ChildObjects.Insert(0, start);
document.Sections.get(0).Paragraphs.get_Item(0).ChildObjects.Add(end);
```

---

# Word Document Protection
## Protect a Word document with specified protection type and password
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Protect the Word file
document.Protect({type: wasmModule.ProtectionType.AllowOnlyReading, password: "123456"});

// Define the output file name
const outputFileName = "SpecifiedProtectionType.docx";

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

// Clean up resources
document.Dispose();
```

---

# Word to PDF Encryption
## Convert Word document to encrypted PDF with password protection
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Create an instance of ToPdfParameterList
let toPdf = wasmModule.ToPdfParameterList.Create();

// Set the user password for the resulted PDF file
toPdf.PdfSecurity.Encrypt("e-iceblue");

// Define the output file name
const outputFileName = "WordToPdfEncrypt.pdf";

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName, paramList: toPdf});

// Clean up resources
document.Dispose();
```

---

# spire.doc javascript field
## add TC field in Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Create a new paragraph
let paragraph = section.AddParagraph();

// Add TC field in the paragraph
let field = paragraph.AppendField("TC", wasmModule.FieldType.FieldTOCEntry);
field.Code = "TC " + "\"Entry Text\"" + " \\f" + " t";
```

---

# spire.doc javascript field conversion
## convert form fields to body text in word document
```javascript
// Create a new document
const doc = wasmModule.Document.Create();

// Load a document from the virtual file system
doc.LoadFromFile(inputFileName);

// Traverse FormFields
for (let i = 0; i < doc.Sections.get(0).Body.FormFields.Count; i++) {
  let field = doc.Sections.get(0).Body.FormFields.get_Item(i);

  // Find FieldFormTextInput type field
  if (field.Type == wasmModule.FieldType.FieldFormTextInput) {
    // Get the paragraph
    let paragraph = field.OwnerParagraph;

    // Define variables
    let startIndex = 0;
    let endIndex = 0;

    // Create a new TextRange
    let textRange = wasmModule.TextRange.Create(doc);

    // Set text for textRange
    textRange.Text = paragraph.Text;

    // Traverse DocumentObjectS of field paragraph
    for (let j = 0; j < paragraph.ChildObjects.Count; j++) {
      let obj = paragraph.ChildObjects.get(j);
      // If its DocumentObjectType is BookmarkStart
      if (obj.DocumentObjectType === wasmModule.DocumentObjectType.BookmarkStart) {
        // Get the index
        startIndex = paragraph.ChildObjects.IndexOf(obj);
      }
      // If its DocumentObjectType is BookmarkEnd
      if (obj.DocumentObjectType === wasmModule.DocumentObjectType.BookmarkEnd) {
        // Get the index
        endIndex = paragraph.ChildObjects.IndexOf(obj);
      }
    }
    // Remove ChildObjects
    for (let k = endIndex; k > startIndex; k--) {
      // If it is TextFormField
      if (paragraph.ChildObjects.get(k) instanceof wasmModule.TextFormField) {
        let textFormField = paragraph.ChildObjects.get(k);

        // Remove the field object
        paragraph.ChildObjects.Remove(textFormField);
      } else {
        paragraph.ChildObjects.RemoveAt(k);
      }
    }
    // Insert the new TextRange
    paragraph.ChildObjects.Insert(startIndex, textRange);
    break;
  }
}
```

---

# Word Document IF Field Conversion
## Convert IF fields to text while preserving formatting
```javascript
// Get all fields in document
let fields = document.Fields;

// Get the total number of fields
let count = fields.Count;

// Loop through each field in the document
for (let i = 0; i < count; i++) {
  
  // Get the field in the collection
  let field = fields.get_Item(i);
  
  // Check if the field is of type 'FieldIf'
  if (field.Type == wasmModule.FieldType.FieldIf) {
    let original = field;
    
    // Get the text of the field
    let text = field.FieldText;
    
    // Create a new textRange and set its format
    let textRange = wasmModule.TextRange.Create(document);
    textRange.Text = text;
    textRange.CharacterFormat.FontName = original.CharacterFormat.FontName;
    textRange.CharacterFormat.FontSize = original.CharacterFormat.FontSize;

    // Get the paragraph that owns the field
    let par = field.OwnerParagraph;
    
    // Get the index of the field
    let index = par.ChildObjects.IndexOf(field);
    
    // Remove the original field from the paragraph
    par.ChildObjects.RemoveAt(index);
    
    // Insert the new text range at the original field's position
    par.ChildObjects.Insert(index, textRange);
  }
}
```

---

# Spire.Doc JavaScript Cross-Reference
## Create a cross-reference to bookmark in Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Create a bookmark
let paragraph = section.AddParagraph();
paragraph.AppendBookmarkStart("MyBookmark");
paragraph.AppendText("Text inside a bookmark");
paragraph.AppendBookmarkEnd("MyBookmark");

// Insert line breaks
for (let i = 0; i < 4; i++) {
    paragraph.AppendBreak(wasmModule.BreakType.LineBreak);
}

// Create a cross-reference field, and link it to bookmark
let field = wasmModule.Field.Create(document);
field.Type = wasmModule.FieldType.FieldRef;
field.Code = "REF MyBookmark \\p \\h";

// Insert field to paragraph
paragraph = section.AddParagraph();
paragraph.AppendText("For more information, see ");
paragraph.ChildObjects.Add(field);

// Insert FieldSeparator object
let fieldSeparator = wasmModule.FieldMark.Create(document, wasmModule.FieldMarkType.FieldSeparator);
paragraph.ChildObjects.Add(fieldSeparator);

// Set display text of the field
let tr = wasmModule.TextRange.Create(document);
tr.Text = "above";
paragraph.ChildObjects.Add(tr);

// Insert FieldEnd object to mark the end of the field
let fieldEnd = wasmModule.FieldMark.Create(document, wasmModule.FieldMarkType.FieldEnd);
paragraph.ChildObjects.Add(fieldEnd);
```

---

# spire.doc javascript form fields
## create form fields in a word document
```javascript
// Create a new section
let section = document.AddSection();

// Add title
AddTitle(section);

// Add form
AddForm(section, xmlDocument);

// Function to add a title to the section
const AddTitle = (section) => {
  // Create a new paragraph for the title
  let title = section.AddParagraph(); 

  // Append text to the title paragraph
  let titleText = title.AppendText("Create Your Account"); 

  // Set the font size of the title text
  titleText.CharacterFormat.FontSize = 18; 

  // Set the font name of the title text
  titleText.CharacterFormat.FontName = "Arial"; 

  // Set the text color of the title using ARGB format
  titleText.CharacterFormat.TextColor = wasmModule.Color.FromArgb(0x00, 0x71, 0xb6); 

  // Align the title paragraph to the center
  title.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center; 

  // Set the spacing after the title paragraph
  title.Format.AfterSpacing = 8;
};

// Function to add forms
const AddForm = (section, xmlDocument) => {
  // Create a paragraph style for the description
  let descriptionStyle = wasmModule.ParagraphStyle.Create(section.Document);
  descriptionStyle.Name = "description";
  descriptionStyle.CharacterFormat.FontSize = 12;
  descriptionStyle.CharacterFormat.FontName = "Arial";
  descriptionStyle.CharacterFormat.TextColor = wasmModule.Color.FromArgb(0x00, 0x45, 0x8e);
  section.Document.Styles.Add(descriptionStyle);

  // Create the first paragraph with instructions
  let p1 = section.AddParagraph();
  let text1
      = "So that we can verify your identity and find your information, "
      + "please provide us with the following information. "
      + "This information will be used to create your online account. "
      + "Your information is not public, shared in anyway, or displayed on this site";
  p1.AppendText(text1);
  p1.ApplyStyle(descriptionStyle.Name);

  // Create the second paragraph with additional instructions
  let p2 = section.AddParagraph();
  let text2 = "You must provide a real email address to which we will send your password.";
  p2.AppendText(text2);
  p2.ApplyStyle(descriptionStyle.Name);
  p2.Format.AfterSpacing = 8;

  // Create a style for form field group labels
  let formFieldGroupLabelStyle = wasmModule.ParagraphStyle.Create(section.Document);
  formFieldGroupLabelStyle.Name = "formFieldGroupLabel";
  formFieldGroupLabelStyle.ApplyBaseStyle("description");
  formFieldGroupLabelStyle.CharacterFormat.Bold = true;
  formFieldGroupLabelStyle.CharacterFormat.TextColor = wasmModule.Color.get_White();
  section.Document.Styles.Add(formFieldGroupLabelStyle);

  // Create a style for form field labels
  let formFieldLabelStyle = wasmModule.ParagraphStyle.Create(section.Document);
  formFieldLabelStyle.Name = "formFieldLabel";
  formFieldLabelStyle.ApplyBaseStyle("description");
  formFieldLabelStyle.ParagraphFormat.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;
  section.Document.Styles.Add(formFieldLabelStyle);

  // Create a table to organize form fields
  let table = section.AddTable();
  table.DefaultColumnsNumber = 2;
  table.DefaultRowHeight = 20;

  let sectionNodes = xmlDocument.documentElement.getElementsByTagName("section");

  // Iterate over each section node to create table rows
  for (let node of sectionNodes) {
    // Add a new row for the section
    let row = table.AddRow(false);
    row.Cells.get(0).CellFormat.BackColor = wasmModule.Color.FromArgb(0x00, 0x71, 0xb6);
    row.Cells.get(0).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;

    // Create a paragraph in the cell
    let cellParagraph = row.Cells.get(0).AddParagraph();

    // Append section name text
    cellParagraph.AppendText(node.getAttribute("name"));

    // Apply the group label style
    cellParagraph.ApplyStyle(formFieldGroupLabelStyle.Name);

    // Select all field nodes within the section
    let fieldNodes = node.querySelectorAll("field");
    for (let fieldNode of fieldNodes) {
      // Add a new row for each field
      let fieldRow = table.AddRow({isCopyFormat: false});
      
      // Center-align the cell content
      fieldRow.Cells.get(0).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
      
      // Create a paragraph in the cell
      let labelParagraph = fieldRow.Cells.get(0).AddParagraph();
      
      // Append the field label text
      labelParagraph.AppendText(fieldNode.getAttribute("label"));
      
      // Apply the field label style
      labelParagraph.ApplyStyle(formFieldLabelStyle.Name);

      // Center-align the cell content
      fieldRow.Cells.get(1).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
      
      // Create a paragraph in the second cell
      let fieldParagraph = fieldRow.Cells.get(1).AddParagraph();
      
      // Get the field ID from the XML
      let fieldId = fieldNode.getAttribute("id");
      
      // Handle different field types based on the XML attributes
      switch (fieldNode.getAttribute("type")) {
            case "text":
              // Append a text input field
              let field = fieldParagraph.AppendField(fieldId, wasmModule.FieldType.FieldFormTextInput);

              field.DefaultText = "";
              field.Text = "";
              break;

            case "list":
              // Append a dropdown list field
              let list = fieldParagraph.AppendField(fieldId, wasmModule.FieldType.FieldFormDropDown);

              let itemNodes = fieldNode.querySelectorAll("item");
              for (let itemNode of itemNodes) {
                  list.DropDownItems.Add(itemNode.textContent);
              }
              break;

            case "checkbox":
              // Append a checkbox field
              fieldParagraph.AppendField(fieldId, wasmModule.FieldType.FieldFormCheckBox);
              break;
          }
      }

      // Merge the first two cells of the row horizontally
      table.ApplyHorizontalMerge(row.GetRowIndex(), 0, 1);
  }
};
```

---

# Creating IF Field in Word Document
## Create and configure an IF field in a Word document using JavaScript
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Create a new paragraph
let paragraph = section.AddParagraph();

// Define a method of creating an IF Field
CreateIfField(document, paragraph);

// Define merged data
let fieldName = ["Count"];
let fieldValue = ["2"];

// Merge data into the IF Field
document.MailMerge.Execute({fieldNames: fieldName, fieldValues: fieldValue});

// Update all fields in the document
document.IsUpdateFields = true;

const CreateIfField = (document, paragraph) => {
  // Create an IF field in the document
  let ifField = wasmModule.IfField.Create(document);
  ifField.Type = wasmModule.FieldType.FieldIf;
  ifField.Code = "IF ";
  
  // Add the IF field to the paragraph
  paragraph.Items.Add(ifField);

  // Append a merge field named "Count" to the paragraph
  paragraph.AppendField("Count", wasmModule.FieldType.FieldMergeField);

  // Append text to the paragraph to complete the IF condition
  paragraph.AppendText(" > ");
  paragraph.AppendText("\"100\" ");
  paragraph.AppendText("\"Thanks\" ");
  paragraph.AppendText("\"The minimum order is 100 units\"");

  // Create an end marker for the IF field
  let end = document.CreateParagraphItem(wasmModule.ParagraphItemType.FieldMark);

  // Add the end marker to the paragraph
  end.Type = wasmModule.FieldMarkType.FieldEnd;
  paragraph.Items.Add(end);

  // Link the end marker to the IF field
  ifField.End = end;
};
```

---

# spire.doc javascript field
## create nested if field in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Create a new paragraph
let paragraph = section.AddParagraph();

// Create an IF field
let ifField = wasmModule.IfField.Create(document);
ifField.Type = wasmModule.FieldType.FieldIf;
ifField.Code = "IF ";
paragraph.Items.Add(ifField);

// Create the embedded IF field
let ifField2 = wasmModule.IfField.Create(document);
ifField2.Type = wasmModule.FieldType.FieldIf;
ifField2.Code = "IF ";
paragraph.ChildObjects.Add(ifField2);
paragraph.Items.Add(ifField2);
paragraph.AppendText("\"200\" < \"50\"   \"200\" \"50\" ");
let embeddedEnd = document.CreateParagraphItem(wasmModule.ParagraphItemType.FieldMark);
embeddedEnd.Type = wasmModule.FieldMarkType.FieldEnd;
paragraph.Items.Add(embeddedEnd);
ifField2.End = embeddedEnd;

paragraph.AppendText(" > ");
paragraph.AppendText("\"100\" ");
paragraph.AppendText("\"Thanks\" ");
paragraph.AppendText("\"The minimum order is 100 units\"");
let end = document.CreateParagraphItem(wasmModule.ParagraphItemType.FieldMark);
end.Type = wasmModule.FieldMarkType.FieldEnd;
paragraph.Items.Add(end);
ifField.End = end;

// Update all fields in the document
document.IsUpdateFields = true;
```

---

# Spire.Doc JavaScript Form Fields
## Fill form fields in a Word document with data from XML
```javascript
// Fill data into form fields
for (let i = 0; i < document.Sections.get(0).Body.FormFields.Count; i++) {
    let field = document.Sections.get(0).Body.FormFields.get_Item(i);
    let path = field.Name;
    let propertyNode = user.querySelector(path);
    
    // Check if the property node exists
    if (propertyNode != null) {
      // Switch based on the type of the form field
      switch (field.Type) {
          case wasmModule.FieldType.FieldFormTextInput:
              // If the field is a text input, set its text to the property node's text content
              field.Text = propertyNode.textContent;
              break;

          case wasmModule.FieldType.FieldFormDropDown:
              // If the field is a dropdown, find the correct item to select
              let combox = field;
              for (let i = 0; i < combox.DropDownItems.Count; i++) {
                // Check if the item text matches the property value
                  if (combox.DropDownItems.get_Item(i).Text === propertyNode.Value) {
                     // Set the selected index
                      combox.DropDownSelectedIndex = i;
                      break;
                  }
                  // Special case for the "country" field to select "Others" if applicable
                  if (field.Name == "country" && combox.DropDownItems.get_Item(i).Text === "Others") {
                      combox.DropDownSelectedIndex = i;
                  }
              }
              break;

            case wasmModule.FieldType.FieldFormCheckBox:
               // If the field is a checkbox, check if it should be checked
              if (propertyNode.textContent) {
                  let checkBox = field;
                  // Set the checkbox to checked
                  checkBox.Checked = true;
              }
              break;
        }
      }   
  }
```

---

# Form Field Properties
## Modify text input form field properties in a Word document
```javascript
// Get the first section
let section = document.Sections.get(0);

// Get FormField by index
let formField = section.Body.FormFields.get_Item(1);

// Check if the form field is of type 'FieldFormTextInput'
if (formField.Type == wasmModule.FieldType.FieldFormTextInput) {
  // Set the text of the form field, incorporating its name
  formField.Text = "My name is " + formField.Name;
  // Set the text color of the form field to red
  formField.CharacterFormat.TextColor = wasmModule.Color.get_Red();
  // Set the text style to italic
  formField.CharacterFormat.Italic = true;
}
```

---

# spire.doc javascript field text
## get field text from document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

let sb = [];

// Get all fields in document
let fields = document.Fields;

for (let i = 0; i < fields.Count; i++) {
    let field = fields.get_Item(i);
    // Get field text
    let fieldText = field.FieldText;
    sb.push("The field text is \"" + fieldText + "\".\r\n");
}

// Combine all the found data into a single string
let content = sb.join("\n");
```

---

# Spire.Doc JavaScript Form Field
## Get form field by name
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Get the first section
let section = document.Sections.get(0);

// Get form field by name
let formField = section.Body.FormFields.get_Item({formFieldName: "email"});
let name = formField.Name;
let formFieldType = formField.FormFieldType;
```

---

# Get Form Fields Collection
## Extract form fields collection from a document section
```javascript
// Get the first section
let section = document.Sections.get(0);

// Get the form fields collection from section
let formFields = section.Body.FormFields;
```

---

# Spire.Doc JavaScript Get Merge Field Names
## Extract merge field names from a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

let sb = [];

// Get merge field name
let fieldNames = document.MailMerge.GetMergeFieldNames();

sb.push("The document has " + fieldNames.length + " merge fields.");
sb.push("The below is the name of the merge field:");
for (let name of fieldNames) {
    sb.push(name);
}

// Combine all the found data into a single string
let content = sb.join("\n");

// Clean up resources
document.Dispose();
```

---

# spire.doc javascript field
## insert address block field in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Add a new paragraph
let par = section.AddParagraph();

// Add address block in the paragraph
let field = par.AppendField("ADDRESSBLOCK", wasmModule.FieldType.FieldAddressBlock);

// Set field code
field.Code = "ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\"";
```

---

# spire.doc javascript advanced field
## insert advance field in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Add a new paragraph
let par = section.AddParagraph();

// Add advance field
let field = par.AppendField("Field", wasmModule.FieldType.FieldAdvance);

// Add field code
field.Code = "ADVANCE \\d 10 \\l 10 \\r 10 \\u 0 \\x 100 \\y 100 ";

// Update field
document.IsUpdateFields = true;
```

---

# Spire.Doc JavaScript Merge Field
## Insert merge field in Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Add a new paragraph
let par = section.AddParagraph();

// Add merge field in the paragraph
par.AppendField("MyFieldName", wasmModule.FieldType.FieldMergeField);
```

---

# Insert None Field in Word Document
## This code demonstrates how to insert a none field into a Word document using Spire.Doc for JavaScript.
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Add a new paragraph
let par = section.AddParagraph();

// Add a none field in the paragraph
par.AppendField("", wasmModule.FieldType.FieldNone);
```

---

# spire.doc javascript page reference field
## insert page reference field in word document
```javascript
// Create a new section
let section = document.LastSection;

// Add a new paragraph
let par = section.AddParagraph();

// Add a page ref field in the paragraph
let field = par.AppendField("pageRef", wasmModule.FieldType.FieldPageRef);

// Set field code
field.Code = "PAGEREF  bookmark1 \\# \"0\" \\* Arabic  \\* MERGEFORMAT";
```

---

# Word Document Custom Property Fields Removal
## Remove all custom property fields from a Word document
```javascript
// Get custom document properties object
let cdp = document.CustomDocumentProperties;

// Remove all custom property fields in the document
for (let i = 0; i < cdp.Count;) {
    cdp.Remove(cdp.get_Item({index: i}).Name);
}

document.IsUpdateFields = true;
```

---

# spire.doc javascript field
## remove field from word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Get the first field
let field = document.Fields.get_Item(0);

// Get the paragraph of the field
let par = field.OwnerParagraph;

// Get the index of the field
let index = par.ChildObjects.IndexOf(field);

// Remove field via index
par.ChildObjects.RemoveAt(index);
```

---

# Spire.Doc JavaScript Text Replacement
## Replace text with merge field in a Word document
```javascript
// Find the text that will be replaced
let ts = document.FindString("Test", true, true);

// Get the text range of the found string
let tr = ts.GetAsOneRange();

// Access the paragraph that owns the text range
let par = tr.OwnerParagraph;

// Get the index of the text range in the paragraph
let index = par.ChildObjects.IndexOf(tr);

// Create a new merge field in the document
let field = wasmModule.MergeField.Create(document);
field.FieldName = "MergeField";

// Insert the new merge field at the specific position in the paragraph
par.ChildObjects.Insert(index, field);

// Remove the original text
par.ChildObjects.Remove(tr);
```

---

# spire.doc javascript fields
## update fields in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Update fields
document.IsUpdateFields = true;
```

---

# spire.doc javascript TOC style
## change table of contents style in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Define a Toc style
let tocStyle = wasmModule.Style.CreateBuiltinStyle({bStyle: wasmModule.BuiltinStyle.Toc1, doc: document});
tocStyle.CharacterFormat.FontName = "Aleo";
tocStyle.CharacterFormat.FontSize = 15;
tocStyle.CharacterFormat.TextColor = wasmModule.Color.get_CadetBlue();
document.Styles.Add(tocStyle);

// Loop through sections
for (let i = 0; i < document.Sections.Count; i++) {
    let section = document.Sections.get(i);
    // Loop through content of section
    for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
        let obj = section.Body.ChildObjects.get(j);
        // Find the structure document tag
        if (obj instanceof wasmModule.StructureDocumentTag) {
            let tag = obj;
            // Find the paragraph where the TOC1 locates
            for (let k = 0; k < tag.ChildObjects.Count; k++) {
                let cObj = tag.ChildObjects.get(k);
                if (cObj instanceof wasmModule.Paragraph) {
                    let para = cObj;
                    if (para.StyleName == "TOC1") {
                        // Apply the new style for TOC1 paragraph
                        para.ApplyStyle(tocStyle.Name);
                    }
                }
            }
        }
    }
}

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});
```

---

# Spire.Doc JavaScript TOC Tab Style
## Change table of contents tab style in Word document
```javascript
// Loop through sections
for (let i = 0; i < document.Sections.Count; i++) {
    let section = document.Sections.get_Item(i);
    // Loop through content of section
    for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
        let obj = section.Body.ChildObjects.get(j);
        // Find the structure document tag
        if (obj instanceof wasmModule.StructureDocumentTag) {
            let tag = obj;
            // Find the paragraph where the TOC1 locates
            for (let k = 0; k < tag.ChildObjects.Count; k++) {
                let cObj = tag.ChildObjects.get(k);
                if (cObj instanceof wasmModule.Paragraph) {
                    let para = cObj;
                    if (para.StyleName == "TOC2") {
                        // Set the tab style of paragraph
                        for (let a = 0; a < para.Format.Tabs.Count; a++) {
                            let tab = para.Format.Tabs.get_Item(a);
                            tab.Position = tab.Position + 20;
                            tab.TabLeader = wasmModule.TabLeader.NoLeader;
                        }
                    }
                }
            }
        }
    }
}
```

---

# spire.doc javascript table of contents
## create table of contents with default settings in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Add a new section
let section = document.AddSection();

// Add a new paragraph
let para = section.AddParagraph();

// Create a table of contents with default switches (\o "1-3" \h \z)
para.AppendTOC(1, 3);

// Add another paragraph for the title
let par = section.AddParagraph();
let tr = par.AppendText("Flowers");
tr.CharacterFormat.FontSize = 30;
par.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

// Create a new paragraph and set the heading level
let para1 = section.AddParagraph();
para1.AppendText("Ornithogalum");
// Apply the Heading1 style to the paragraph
para1.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading1});
para1 = section.AddParagraph();
// Append a picture to the paragraph
let picture = para1.AppendPicture({imgFile: inputFileName});
// Set the text wrapping style
picture.TextWrappingStyle = wasmModule.TextWrappingStyle.Square;
para1.AppendText("Ornithogalum is a genus of perennial plants mostly native to southern Europe and southern Africa belonging to the family Asparagaceae. Some species are native to other areas such as the Caucasus. Growing from a bulb, species have linear basal leaves and a slender stalk, up to 30 cm tall, bearing clusters of typically white star-shaped flowers, often striped with green.");
para1 = section.AddParagraph();

// Create another paragraph for the next heading
let para2 = section.AddParagraph();
para2.AppendText("Rosa");
para2.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading2});
para2 = section.AddParagraph();
let picture2 = para2.AppendPicture({imgFile: inputFileName1});
picture2.TextWrappingStyle = wasmModule.TextWrappingStyle.Square;
para2.AppendText("A rose is a woody perennial flowering plant of the genus Rosa, in the family Rosaceae, or the flower it bears. There are over a hundred species and thousands of cultivars. They form a group of plants that can be erect shrubs, climbing or trailing with stems that are often armed with sharp prickles. Flowers vary in size and shape and are usually large and showy, in colours ranging from white through yellows and reds. Most species are native to Asia, with smaller numbers native to Europe, North America, and northwestern Africa. Species, cultivars and hybrids are all widely grown for their beauty and often are fragrant. Roses have acquired cultural significance in many societies. Rose plants range in size from compact, miniature roses, to climbers that can reach seven meters in height. Different species hybridize easily, and this has been used in the development of the wide range of garden roses.");
section.AddParagraph();

// Create another paragraph for the next heading
let para3 = section.AddParagraph();
para3.AppendText("Hyacinth");
para3.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading3});
para3 = section.AddParagraph();
let picture3 = para3.AppendPicture({imgFile: inputFileName2});
picture3.TextWrappingStyle = wasmModule.TextWrappingStyle.Tight;
para3.AppendText("Hyacinthus is a small genus of bulbous, fragrant flowering plants in the family Asparagaceae, subfamily Scilloideae.These are commonly called hyacinths.The genus is native to the eastern Mediterranean (from the south of Turkey through to northern Israel).");
para3 = section.AddParagraph();
para3.AppendText("Several species of Brodiea, Scilla, and other plants that were formerly classified in the lily family and have flower clusters borne along the stalk also have common names with the word \"hyacinth\" in them. Hyacinths should also not be confused with the genus Muscari, which are commonly known as grape hyacinths.");

// Update TOC
document.UpdateTableOfContents();
```

---

# spire.doc javascript table of contents
## customize table of contents in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Add a new section
let section = document.AddSection();

// Customize the table of contents with specific switches
let toc = wasmModule.TableOfContent.Create(document, "{\\o \"1-3\" \\n 1-1}");

// Add a new paragraph to the section
let para = section.AddParagraph();

// Add the table of contents to the paragraph
para.Items.Add(toc);

// Append a field mark separator for the TOC
para.AppendFieldMark(wasmModule.FieldMarkType.FieldSeparator);

// Append the text "TOC" to indicate the table of contents
para.AppendText("TOC");

// Append the end field mark for the TOC
para.AppendFieldMark(wasmModule.FieldMarkType.FieldEnd);

// Set the document's TOC to the newly created TOC
document.TOC = toc;

// Add a new paragraph for the title
let par = section.AddParagraph();
let tr = par.AppendText("Flowers");
tr.CharacterFormat.FontSize = 30;
par.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

// Create a paragraph for the first heading
let para1 = section.AddParagraph();
para1.AppendText("Ornithogalum");
para1.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading1});

// Create a paragraph for the next heading
let para2 = section.AddParagraph();
para2.AppendText("Rosa");
para2.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading2});

// Create a paragraph for the next heading
let para3 = section.AddParagraph();
para3.AppendText("Hyacinth");
para3.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading3});

// Update TOC
document.UpdateTableOfContents();
```

---

# spire.doc javascript table of content
## remove table of content from Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Get the first body from the first section
let body = document.Sections.get(0).Body;

// Remove TOC from first body
let regex = wasmModule.Regex.Create("TOC\\w+", wasmModule.RegexOptions.None);
for (let i = 0; i < body.Paragraphs.Count; i++) {
    if (regex.IsMatch(body.Paragraphs.get_Item(i).StyleName)) {
        body.Paragraphs.RemoveAt(i);
        i--;
    }
}

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

// Clean up resources
document.Dispose();
```

---

# spire.doc javascript textbox
## delete table from textbox
```javascript
// Get the first textbox
let textbox = document.TextBoxes.get_Item(0);

// Remove the first table from the textbox
textbox.Body.Tables.RemoveAt(0);
```

---

# Extract Text from TextBoxes in Word Document
## This code demonstrates how to extract text from textboxes in a Word document, including text within paragraphs and tables inside textboxes.
```javascript
let sw = [];

// Verify whether the document contains a textbox or not.
if (document.TextBoxes.Count > 0) {
    // Traverse the document.
    for (let i = 0; i < document.Sections.Count; i++) {
        let section = document.Sections.get(i);
        for (let j = 0; j < section.Paragraphs.Count; j++) {
            let p = section.Paragraphs.get_Item(j);
            for (let k = 0; k < p.ChildObjects.Count; k++) {
                let obj = p.ChildObjects.get(k);
                if (obj.DocumentObjectType == wasmModule.DocumentObjectType.TextBox) {
                    let textbox = obj;
                    for (let a = 0; a < textbox.ChildObjects.Count; a++) {
                        let objt = textbox.ChildObjects.get(a);
                        //Extract text from paragraph in TextBox.
                        if (objt.DocumentObjectType == wasmModule.DocumentObjectType.Paragraph) {
                            sw.push(objt.Text);
                        }

                        // Extract text from Table in TextBox.
                        if (objt.DocumentObjectType == wasmModule.DocumentObjectType.Table) {
                            let table = objt;
                            ExtractTextFromTables(table, sw);
                        }
                    }
                }
            }
        }
    }
}

const ExtractTextFromTables = (table, sw) => {
    for (let i = 0; i < table.Rows.Count; i++) {
        let row = table.Rows.get(i);
        for (let j = 0; j < row.Cells.Count; j++) {
            let cell = row.Cells.get(j);
            for (let k = 0; k < cell.Paragraphs.Count; k++) {
                let paragraph = cell.Paragraphs.get_Item(k);
                sw.push(paragraph.Text);
            }
        }
    }
};
```

---

# spire.doc javascript textbox
## insert image into textbox
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Create a new paragraph
let paragraph = section.AddParagraph();

// Append a textbox to paragraph
let tb = paragraph.AppendTextBox(220, 220);

// Set the position of the textbox
tb.Format.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
tb.Format.HorizontalPosition = 50;
tb.Format.VerticalOrigin = wasmModule.VerticalOrigin.Page;
tb.Format.VerticalPosition = 50;

// Set the fill effect of textbox as picture
tb.Format.FillEfects.Type = wasmModule.BackgroundType.Picture;

// Fill the textbox with a picture
tb.Format.FillEfects.SetPicture(inputFileName);
```

---

# Spire.Doc JavaScript Textbox Table
## Insert a table into a textbox in a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Add a paragraph to the section
let paragraph = section.AddParagraph();

// Add a textbox to the paragraph
let textbox = paragraph.AppendTextBox(300, 100);

// Set the position of the textbox
textbox.Format.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
textbox.Format.HorizontalPosition = 140;
textbox.Format.VerticalOrigin = wasmModule.VerticalOrigin.Page;
textbox.Format.VerticalPosition = 50;

// Add text to the textbox
let textboxParagraph = textbox.Body.AddParagraph();
let textboxRange = textboxParagraph.AppendText("Table 1");
textboxRange.CharacterFormat.FontName = "Arial";

// Insert table to the textbox
let table = textbox.Body.AddTable({showBorder: true});

// Specify the number of rows and columns of the table
table.ResetCells(4, 4);

let data =
    [
        ["Name", "Age", "Gender", "ID"],
        ["John", "28", "Male", "0023"],
        ["Steve", "30", "Male", "0024"],
        ["Lucy", "26", "female", "0025"]
    ];

// Add data to the table
for (let i = 0; i < 4; i++) {
    for (let j = 0; j < 4; j++) {
        let tableRange = table.Rows.get(i).Cells.get(j).AddParagraph().AppendText(data[i][j]);
        tableRange.CharacterFormat.FontName = "Arial";
    }
}

// Apply style to the table
table.ApplyStyle(wasmModule.DefaultTableStyle.TableColorful2);
```

---

# spire.doc javascript textbox
## read table from textbox in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Get the first textbox
let textbox = document.TextBoxes.get_Item(0);

// Get the first table in the textbox
let table = textbox.Body.Tables.get_Item(0);

let str = "";

// Loop through the paragraphs of the table cells and extract them to a .txt file
for (let i = 0; i < table.Rows.Count; i++) {
    let row = table.Rows.get_Item(i);
    for (let j = 0; j < row.Cells.Count; j++) {
        let cell = row.Cells.get_Item(j);
        for (let k = 0; k < cell.Paragraphs.Count; k++) {
            let paragraph = cell.Paragraphs.get_Item(k);
            str += paragraph.Text + "\t";
        }
    }
    str += "\r\n";
}
```

---

# Spire.Doc JavaScript Text Box
## Remove text box from Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document
document.LoadFromFile("TextBoxTemplate.docx");

// Remove the first text box
document.TextBoxes.RemoveAt(0);

// Clean up resources
document.Dispose();
```

---

# spire.doc javascript textbox
## create and format textboxes in word document
```javascript
const InsertTextbox = (section) => {
  let paragraph = section.Paragraphs.Count > 0 ? section.Paragraphs[0] : section.AddParagraph();
  paragraph = section.AddParagraph();
  paragraph = section.AddParagraph();

  // Insert and format the first textbox.
  let textBox1 = paragraph.AppendTextBox(240, 35);
  textBox1.Format.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
  textBox1.Format.LineColor = wasmModule.Color.get_Gray();
  textBox1.Format.LineStyle = wasmModule.TextBoxLineStyle.Simple;
  textBox1.Format.FillColor = wasmModule.Color.get_DarkSeaGreen();
  let para = textBox1.Body.AddParagraph();
  let txtrg = para.AppendText("Textbox 1 in the document");
  txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
  txtrg.CharacterFormat.FontSize = 14;
  txtrg.CharacterFormat.TextColor = wasmModule.Color.get_White();
  para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

  // Insert and format the second textbox.
  paragraph = section.AddParagraph();
  paragraph = section.AddParagraph();
  paragraph = section.AddParagraph();
  paragraph = section.AddParagraph();
  let textBox2 = paragraph.AppendTextBox(240, 35);
  textBox2.Format.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
  textBox2.Format.LineColor = wasmModule.Color.get_Tomato();
  textBox2.Format.LineStyle = wasmModule.TextBoxLineStyle.ThinThick;
  textBox2.Format.FillColor = wasmModule.Color.get_Blue();
  textBox2.Format.LineDashing = wasmModule.LineDashing.Dot;
  para = textBox2.Body.AddParagraph();
  txtrg = para.AppendText("Textbox 2 in the document");
  txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
  txtrg.CharacterFormat.FontSize = 14;
  txtrg.CharacterFormat.TextColor = wasmModule.Color.get_Pink();
  para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

  // Insert and format the third textbox.
  paragraph = section.AddParagraph();
  paragraph = section.AddParagraph();
  paragraph = section.AddParagraph();
  paragraph = section.AddParagraph();
  let textBox3 = paragraph.AppendTextBox(240, 35);
  textBox3.Format.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
  textBox3.Format.LineColor = wasmModule.Color.get_Violet();
  textBox3.Format.LineStyle = wasmModule.TextBoxLineStyle.Triple;
  textBox3.Format.FillColor = wasmModule.Color.get_Pink();
  textBox3.Format.LineDashing = wasmModule.LineDashing.DashDotDot;
  para = textBox3.Body.AddParagraph();
  txtrg = para.AppendText("Textbox 3 in the document");
  txtrg.CharacterFormat.FontName = "Lucida Sans Unicode";
  txtrg.CharacterFormat.FontSize = 14;
  txtrg.CharacterFormat.TextColor = wasmModule.Color.get_Tomato();
  para.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
}
```

---

# spire.doc javascript textbox formatting
## set textbox position, line style and internal margin in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a section
let section = document.AddSection();

// Add a text box and append sample text
let TB = section.AddParagraph().AppendTextBox(310, 90);
let para = TB.Body.AddParagraph();
let TR = para.AppendText("Using Spire.Doc, developers will find " +
    "a simple and effective method to endow their applications with rich MS Word features. ");
TR.CharacterFormat.FontName = "Cambria ";
TR.CharacterFormat.FontSize = 13;

// Set exact position for the text box
TB.Format.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
TB.Format.HorizontalPosition = 120;
TB.Format.VerticalOrigin = wasmModule.VerticalOrigin.Page;
TB.Format.VerticalPosition = 100;

// Set line style for the text box
TB.Format.LineStyle = wasmModule.TextBoxLineStyle.Double;
TB.Format.LineColor = wasmModule.Color.get_CornflowerBlue();
TB.Format.LineDashing = wasmModule.LineDashing.Solid;
TB.Format.LineWidth = 5;

// Set internal margin for the text box
TB.Format.InternalMargin.Top = 15;
TB.Format.InternalMargin.Bottom = 10;
TB.Format.InternalMargin.Left = 12;
TB.Format.InternalMargin.Right = 10;
```

---

# Word Document Image Watermark
## Add an image watermark to a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Create a new PictureWatermark instance
let picture = wasmModule.PictureWatermark.Create();

// Set the picture
picture.SetPicture(imageFileName);

// Set the scaling factor for the watermark
picture.Scaling = 250;

// Set the washout property to false (indicating the watermark is not faded)
picture.IsWashout = false;

// Assign the created picture as the watermark for the document
document.Watermark = picture;
```

---

# Spire.Doc JavaScript Watermark Removal
## Remove text and image watermarks from Word documents
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Set the watermark as null to remove the text and image watermark
document.Watermark = null;
```

---

# Spire.Doc JavaScript Text Watermark
## Add text watermark to Word document
```javascript
// Get the first section
let section = document.Sections.get_Item(0);

// Create a new TextWatermark instance
let txtWatermark = wasmModule.TextWatermark.Create();

// Set the text for the watermark
txtWatermark.Text = "E-iceblue";

// Set the font size for the watermark text
txtWatermark.FontSize = 95;

// Set the color of the watermark text to blue
txtWatermark.Color = wasmModule.Color.get_Blue();

// Set the layout of the watermark to diagonal
txtWatermark.Layout = wasmModule.WatermarkLayout.Diagonal;

// Assign the created text watermark to the document
section.Document.Watermark = txtWatermark;
```

---

# Spire.Doc JavaScript OLE Extraction
## Extract OLE objects from Word documents and save them as separate files
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Traverse through all sections of the word document
for (let i = 0; i < document.Sections.Count; i++) {
    let sec = document.Sections.get(i);
    // Traverse through all Child Objects in the body of each section
    for (let j = 0; j < sec.Body.ChildObjects.Count; j++) {
        let obj = sec.Body.ChildObjects.get(j);
        // Find the paragraph
        if (obj instanceof wasmModule.Paragraph) {
            let par = obj;
            for (let k = 0; k < par.ChildObjects.Count; k++) {
                let o = par.ChildObjects.get(k);
                // Check whether the object is OLE
                if (o.DocumentObjectType === wasmModule.DocumentObjectType.OleObject) {
                    let Ole = o;
                    let s = Ole.ObjectType;

                    // Check whether the object type is "Acrobat.Document.11"
                    if (s === "Acrobat.Document.DC") {
                        // Write the data of OLE into file
                        wasmModule.FS.writeFile(PdfOutputFileName, Ole.NativeData);
                    }

                    // Check whether the object type is "Excel.Sheet.8"
                    else if (s === "Excel.Sheet.8") {
                        // Write the data of OLE into file
                        wasmModule.FS.writeFile(XlsOutputFileName, Ole.NativeData);
                    }

                    //  Check whether the object type is "PowerPoint.Show.12"
                    else if (s === "PowerPoint.Show.12") {
                        // Write the data of OLE into file
                        wasmModule.FS.writeFile(PPTOutputFileName, Ole.NativeData);
                    }
                  }
              }
          }
      }
  }

// Clean up resources
document.Dispose();
```

---

# spire.doc javascript OLE
## insert OLE object into Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Add a section
let sec = document.AddSection();

// Add a paragraph
let par = sec.AddParagraph();

// Load the image
let picture = wasmModule.DocPicture.Create(document);
picture.LoadImage({imgFile: imageFileName});

// Insert the OLE
let obj = par.AppendOleObject({
    pathToFile: inputFileName,
    olePicture: picture,
    type: wasmModule.OleObjectType.ExcelWorksheet
});
```

---

# spire.doc javascript ole
## insert OLE object as icon via stream
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Add a section
let sec = document.AddSection();

// Add a paragraph
let par = sec.AddParagraph();

// Create an OLE stream from the specified input file
let stream = wasmModule.Stream.CreateByFile(inputFileName);

// Load the image into a DocPicture object
let picture = wasmModule.DocPicture.Create(document);
picture.LoadImage(imageFileName);

// Insert the OLE object using the created stream and picture
let obj = par.AppendOleObject({
    oleStream: stream,       
    olePicture: picture,      
    fileExtension: "zip"      
});

// Display the OLE object as an icon instead of the content
obj.DisplayAsIcon = true;
```

---

# spire.doc javascript checkbox content control
## add checkbox content control to word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Add a new section
let section = document.AddSection();

// Add a paragraph
let paragraph = section.AddParagraph();

// Append textRange for the paragraph
let txtRange = paragraph.AppendText("The following example shows how to add CheckBox content control in a Word document. \n");

// Append textRange
txtRange = paragraph.AppendText("Add CheckBox Content Control:  ");

// Set the font format
txtRange.CharacterFormat.Italic = true;

// Create StructureDocumentTagInline for document
let sdt = wasmModule.StructureDocumentTagInline.Create(document);

// Add sdt in paragraph
paragraph.ChildObjects.Add(sdt);

// Specify the type
sdt.SDTProperties.SDTType = wasmModule.SdtType.CheckBox;

// Set properties for control
let scb = wasmModule.SdtCheckBox.Create();
sdt.SDTProperties.ControlProperties = scb;

// Add textRange format
let tr = wasmModule.TextRange.Create(document);
tr.CharacterFormat.FontName = "MS Gothic";
tr.CharacterFormat.FontSize = 12;

// Add textRange to StructureDocumentTagInline
sdt.ChildObjects.Add(tr);

// Set checkBox as checked
scb.Checked = true;
```

---

# Spire.Doc JavaScript Content Controls
## Add various types of content controls to a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Add a new section
let section = document.AddSection();

// Add a paragraph
let paragraph = section.AddParagraph();

// Append textRange in the paragraph
let txtRange = paragraph.AppendText("The following example shows how to add content controls in a Word document.");
paragraph = section.AddParagraph();

// Add Combo Box Content Control
paragraph = section.AddParagraph();
txtRange = paragraph.AppendText("Add Combo Box Content Control:  ");
txtRange.CharacterFormat.Italic = true;
let sd = wasmModule.StructureDocumentTagInline.Create(document);
paragraph.ChildObjects.Add(sd);
sd.SDTProperties.SDTType = wasmModule.SdtType.ComboBox;
let cb = wasmModule.SdtComboBox.Create();
cb.ListItems.Add(wasmModule.SdtListItem.Create("Spire.Doc"));
cb.ListItems.Add(wasmModule.SdtListItem.Create("Spire.XLS"));
cb.ListItems.Add(wasmModule.SdtListItem.Create("Spire.PDF"));
sd.SDTProperties.ControlProperties = cb;
let rt = wasmModule.TextRange.Create(document);
rt.Text = cb.ListItems.get_Item(0).DisplayText;
sd.SDTContent.ChildObjects.Add(rt);

section.AddParagraph();

// Add Text Content Control
paragraph = section.AddParagraph();
txtRange = paragraph.AppendText("Add Text Content Control:  ");
txtRange.CharacterFormat.Italic = true;
sd = wasmModule.StructureDocumentTagInline.Create(document);
paragraph.ChildObjects.Add(sd);
sd.SDTProperties.SDTType = wasmModule.SdtType.Text;
let text = wasmModule.SdtText.Create(true);
text.IsMultiline = true;
sd.SDTProperties.ControlProperties = text;
rt = wasmModule.TextRange.Create(document);
rt.Text = "Text";
sd.SDTContent.ChildObjects.Add(rt);

section.AddParagraph();

// Add Picture Content Control
paragraph = section.AddParagraph();
txtRange = paragraph.AppendText("Add Picture Content Control:  ");
txtRange.CharacterFormat.Italic = true;
sd = wasmModule.StructureDocumentTagInline.Create(document);
paragraph.ChildObjects.Add(sd);
sd.SDTProperties.SDTType = wasmModule.SdtType.Picture;
let pic = wasmModule.DocPicture.Create(document);
pic.Width = 10;
pic.Height = 10;
sd.SDTContent.ChildObjects.Add(pic);

section.AddParagraph();

// Add Date Picker Content Control
paragraph = section.AddParagraph();
txtRange = paragraph.AppendText("Add Date Picker Content Control:  ");
txtRange.CharacterFormat.Italic = true;
sd = wasmModule.StructureDocumentTagInline.Create(document);
paragraph.ChildObjects.Add(sd);
sd.SDTProperties.SDTType = wasmModule.SdtType.DatePicker;
let date = wasmModule.SdtDate.Create();
date.CalendarType = wasmModule.CalendarType.Default;
date.DateFormat = "yyyy.MM.dd";
date.FullDate = wasmModule.DateTime.get_Now();
sd.SDTProperties.ControlProperties = date;
rt = wasmModule.TextRange.Create(document);
rt.Text = "1990.02.08";
sd.SDTContent.ChildObjects.Add(rt);

section.AddParagraph();

// Add Drop-Down List Content Control
paragraph = section.AddParagraph();
txtRange = paragraph.AppendText("Add Drop-Down List Content Control:  ");
txtRange.CharacterFormat.Italic = true;
sd = wasmModule.StructureDocumentTagInline.Create(document);
paragraph.ChildObjects.Add(sd);
sd.SDTProperties.SDTType = wasmModule.SdtType.DropDownList;
let sddl = wasmModule.SdtDropDownList.Create();
sddl.ListItems.Add(wasmModule.SdtListItem.Create("Harry"));
sddl.ListItems.Add(wasmModule.SdtListItem.Create("Jerry"));
sd.SDTProperties.ControlProperties = sddl;
rt = wasmModule.TextRange.Create(document);
rt.Text = sddl.ListItems.get_Item(0).DisplayText;
sd.SDTContent.ChildObjects.Add(rt);
```

---

# spire.doc javascript richtext content control
## add richtext content control to word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Add a new section 
let section = document.AddSection();

// Add a paragraph
let paragraph = section.AddParagraph();

// Append textRange for the paragraph
let txtRange = paragraph.AppendText("The following example shows how to add RichText content control in a Word document. \n");

// Append textRange
txtRange = paragraph.AppendText("Add RichText Content Control:  ");

// Set the font format
txtRange.CharacterFormat.Italic = true;

// Create StructureDocumentTagInline for document
let sdt = wasmModule.StructureDocumentTagInline.Create(document);

// Add sdt in paragraph
paragraph.ChildObjects.Add(sdt);

// Specify the type
sdt.SDTProperties.SDTType = wasmModule.SdtType.RichText;

// Set displaying text
let text = wasmModule.SdtText.Create(true);
text.IsMultiline = true;
sdt.SDTProperties.ControlProperties = text;

// Create a TextRange
let rt = wasmModule.TextRange.Create(document);
rt.Text = "Welcome to use ";
rt.CharacterFormat.TextColor = wasmModule.Color.get_Green();
sdt.SDTContent.ChildObjects.Add(rt);

rt = wasmModule.TextRange.Create(document);
rt.Text = "Spire.Doc";
rt.CharacterFormat.TextColor = wasmModule.Color.get_OrangeRed();
sdt.SDTContent.ChildObjects.Add(rt);
```

---

# spire.doc javascript combobox
## add, select and remove combo box item in Word document
```javascript
//Get the combo box from the file
for (let i = 0; i < document.Sections.Count; i++) {
    let section = document.Sections.get_Item(i);
    for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
        let bodyObj = section.Body.ChildObjects.get(j);
        if (bodyObj.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTag) {
            //If SDTType is ComboBox
            if (bodyObj.SDTProperties.SDTType === wasmModule.SdtType.ComboBox) {
                let combo = bodyObj.SDTProperties.ControlProperties;
                //Remove the second list item
                combo.ListItems.RemoveAt(1);
                //Add a new item
                let item = wasmModule.SdtListItem.Create("D", "D");
                combo.ListItems.Add(item);

                //If the value of list items is "D"
                for (let i = 0; i < combo.ListItems.Count; i++) {
                    let sdtItem = combo.ListItems.get_Item(i);
                    if (sdtItem.Value === "D") {
                        //Select the item
                        combo.ListItems.SelectedValue = sdtItem;
                    }
                }
            }
        }
    }
}
```

---

# Get Content Control Properties
## Extract properties from structured document tags in a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

//Get all structureTags in the Word document
let structureTags = GetAllTags(document);

//Get all StructureDocumentTagInline objects
let tagInlines = structureTags.tagInlines;

let property = "";
property += "Alias of contentControl" + "\t" + "ID          " + "\t" + "Tag             " + "\t" + "STDType        " + "\r\n";
//Get properties of all tagInlines
for (let i = 0; i < tagInlines.length; i++) {
    let alias = tagInlines[i].SDTProperties.Alias;
    let id = tagInlines[i].SDTProperties.Id;
    let tag = tagInlines[i].SDTProperties.Tag;
    let STDType = tagInlines[i].SDTProperties.SDTType.toString();
    property += alias + ",\t" + id + ",\t" + tag + ",\t" + STDType + "\r\n";
}

//Get all StructureDocumentTag objects
let tags = structureTags.tags;

//Get properties of all tags
for (let i = 0; i < tags.length; i++) {
    let alias = tags[i].SDTProperties.Alias;
    let id = tags[i].SDTProperties.Id;
    let tag = tags[i].SDTProperties.Tag;
    let STDType = tags[i].SDTProperties.SDTType.toString();
    property += alias + ",\t" + id + ",\t" + tag + ",\t" + STDType + "\r\n";
}

const GetAllTags = (document) => {
    let tagInlines = [];
    let tags = [];

    let structureTags = new StructureTags(tagInlines, tags);
    for (let i = 0; i < document.Sections.Count; i++) {
        let section = document.Sections.get_Item(i);
        for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
            let obj = section.Body.ChildObjects.get(j);
            if (obj.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTag) {
                structureTags.tags.push(obj);
            } else if (obj.DocumentObjectType === wasmModule.DocumentObjectType.Paragraph) {
                for (let j = 0; j < obj.ChildObjects.Count; j++) {
                    let pobj = obj.ChildObjects.get(j);
                    if (pobj.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTagInline) {
                        structureTags.tagInlines.push(pobj);
                    }
                }
            }
            else if (obj.DocumentObjectType === wasmModule.DocumentObjectType.Table) {
                for (let a = 0; a < obj.Rows.Count; a++) {
                    let row = obj.Rows.get_Item(a);
                    for (let b = 0; b < row.Cells.Count; b++) {
                        let cell = row.Cells.get_Item(j);
                        for (let c = 0; c < cell.ChildObjects.Count; c++) {
                            let cellChild = cell.ChildObjects.get(c);
                            if (cellChild.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTag) {
                                structureTags.tags.push(cellChild);
                            } else if (cellChild.DocumentObjectType === wasmModule.DocumentObjectType.Paragraph) {
                                for (let d = 0; d < cellChild.ChildObjects.Count; d++) {
                                    let pobj = cellChild.ChildObjects.get(d);
                                    if (pobj.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTagInline) {
                                        structureTags.tagInlines.push(pobj);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }
    return structureTags;
};

class StructureTags {
    constructor(tagInlines, tags) {
        this.tagInlines = tagInlines;
        this.tags = tags;
    }
}
```

---

# Spire.Doc JavaScript Content Control
## Lock content control content in Word document
```javascript
//Create StructureDocumentTag for document
let sdt = wasmModule.StructureDocumentTag.Create(document);
let section2 = document.AddSection();
section2.Body.ChildObjects.Add(sdt);

//Specify the type
sdt.SDTProperties.SDTType = wasmModule.SdtType.RichText;

for (let i = 0; i < section.Body.ChildObjects.Count; i++) {
    let obj = section.Body.ChildObjects.get(i);
    if (obj.DocumentObjectType === wasmModule.DocumentObjectType.Table) {
        sdt.SDTContent.ChildObjects.Add(obj.Clone());
    }
}

// Lock content
sdt.SDTProperties.LockSettings = wasmModule.LockSettingsType.ContentLocked;

document.Sections.Remove(section);
```

---

# spire.doc javascript content controls
## remove content controls from word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

//Loop through sections
for (let s = 0; s < document.Sections.Count; s++) {
    let section = document.Sections.get(s);
    for (let i = 0; i < section.Body.ChildObjects.Count; i++) {
        //Loop through contents in paragraph
        if (section.Body.ChildObjects.get(i) instanceof wasmModule.Paragraph) {
            let para = section.Body.ChildObjects.get(i);
            for (let j = 0; j < para.ChildObjects.Count; j++) {
                //Find the StructureDocumentTagInline
                if (para.ChildObjects.get(j) instanceof wasmModule.StructureDocumentTagInline) {
                    let sdt = para.ChildObjects.get(j);
                    //Remove the content control from paragraph
                    para.ChildObjects.Remove(sdt);
                    j--;
                }
            }
        }
        if (section.Body.ChildObjects.get(i) instanceof wasmModule.StructureDocumentTag) {
            let sdt = section.Body.ChildObjects.get(i);
            section.Body.ChildObjects.Remove(sdt);
            i--;
        }
    }
}

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

// Clean up resources
document.Dispose();
```

---

# spire.doc javascript structured document tag
## update checkbox content control in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Get all structured document tags
let tagInlines = GetAllTags(document);

// Get the controls
for (let i = 0; i < tagInlines.length; i++) {
    // Get the type
    let type = tagInlines[i].SDTProperties.SDTType;

    // Update the status
    if (type === wasmModule.SdtType.CheckBox) {
        let scb = tagInlines[i].SDTProperties.ControlProperties;
        if (scb.Checked) {
            scb.Checked = false;
        } else {
            scb.Checked = true;
        }
    }
}

// Save the document to the specified path
document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

// Clean up resources
document.Dispose();

const GetAllTags = (document) => {
  let tagInlines = [];
  // Travel document sections
  for (let i = 0; i < document.Sections.Count; i++) {
      let section = document.Sections.get(i);
      for (let j = 0; j < section.Body.ChildObjects.Count; j++) {
          let obj = section.Body.ChildObjects.get(j);
          // Travel document paragraphs
          if (obj.DocumentObjectType === wasmModule.DocumentObjectType.Paragraph) {
              for (let k = 0; k < obj.ChildObjects.Count; k++) {
                  let pobj = obj.ChildObjects.get(k);
                  // Get StructureDocumentTagInline
                  if (pobj.DocumentObjectType === wasmModule.DocumentObjectType.StructureDocumentTagInline) {
                      tagInlines.push(pobj);
                  }
              }
          }
      }
    }
    return tagInlines;
}
```

---

# Spire.Doc JavaScript Endnote
## Insert and format endnote in Word document
```javascript
// Get the first section of the document
let s = document.Sections.get_Item(0);

// Get the second paragraph in the section
let p = s.Paragraphs.get_Item(1);

// Add an endnote to the paragraph
let endnote = p.AppendFootnote({type: wasmModule.FootnoteType.Endnote});

// Append a new paragraph to the endnote's text body and add text
let text = endnote.TextBody.AddParagraph().AppendText("Reference: Wikipedia");

// Set the text format for the endnote content
text.CharacterFormat.FontName = "Impact";
text.CharacterFormat.FontSize = 14;
text.CharacterFormat.TextColor = wasmModule.Color.get_DarkOrange(); 

// Set the marker format for the endnote reference
endnote.MarkerCharacterFormat.FontName = "Calibri"; 
endnote.MarkerCharacterFormat.FontSize = 25; 
endnote.MarkerCharacterFormat.TextColor = wasmModule.Color.get_DarkBlue();
```

---

# Spire.Doc JavaScript Footnote
## Insert footnote into Word document
```javascript
// Find the first matched string in the document
let selection = document.FindString("Spire.Doc", false, true);

// Get the text range of the found string
let textRange = selection.GetAsOneRange();

// Get the paragraph that contains the matched text range
let paragraph = textRange.OwnerParagraph;

// Get the index of the text range within the paragraph's child objects
let index = paragraph.ChildObjects.IndexOf(textRange);

// Append a footnote to the paragraph
let footnote = paragraph.AppendFootnote({ type: wasmModule.FootnoteType.Footnote });

// Insert the footnote into the paragraph just after the matched text range
paragraph.ChildObjects.Insert(index + 1, footnote);

// Add a new paragraph to the footnote's text body and append text
textRange = footnote.TextBody.AddParagraph().AppendText("Welcome to evaluate Spire.Doc");

// Set the text format for the footnote content
textRange.CharacterFormat.FontName = "Arial Black"; 
textRange.CharacterFormat.FontSize = 10; 
textRange.CharacterFormat.TextColor = wasmModule.Color.get_DarkGray(); 

// Set the marker format for the footnote reference
footnote.MarkerCharacterFormat.FontName = "Calibri";
footnote.MarkerCharacterFormat.FontSize = 12; 
footnote.MarkerCharacterFormat.Bold = true; 
footnote.MarkerCharacterFormat.TextColor = wasmModule.Color.get_DarkGreen();
```

---

# Spire.Doc JavaScript Footnote Removal
## Remove footnotes from a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Get the first section of the document
let section = document.Sections.get(0);

// Traverse paragraphs in the section to find and remove footnotes
for (let p = 0; p < section.Paragraphs.Count; p++) {
    let para = section.Paragraphs.get_Item(p);
    let index = -1; 

    // Check each child object in the paragraph to find a footnote
    for (let i = 0, cnt = para.ChildObjects.Count; i < cnt; i++) {
        let pBase = para.ChildObjects.get(i); 
        if (pBase instanceof wasmModule.Footnote) {
            index = i; 
            break; 
        }
    }

    // If a footnote was found, remove it from the paragraph
    if (index > -1) {
        para.ChildObjects.RemoveAt(index); 
    }
}
```

---

# spire.doc javascript footnote
## set footnote position and number format
```javascript
// Get the first section
let sec = document.Sections.get(0);

// Set the number format, restart rule and position for the footnote
sec.FootnoteOptions.NumberFormat = wasmModule.FootnoteNumberFormat.UpperCaseLetter;
sec.FootnoteOptions.RestartRule = wasmModule.FootnoteRestartRule.RestartPage;
sec.FootnoteOptions.Position = wasmModule.FootnotePosition.PrintAsEndOfSection;
```

---

# Spire.Doc JavaScript VBA Macro Handling
## Detect and remove VBA macros from Word documents
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document
document.LoadFromFile(inputFileName);

// If the document contains Macros, remove them from the document
if (document.IsContainMacro) {
    document.ClearMacros();
}

// Clean up resources
document.Dispose();
```

---

# spire.doc javascript macros
## load and save word document with macros
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Define the output file name
const outputFileName = "Macros.docm";

// Save the document to the specified path with macro enabled format
document.SaveToFile({fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docm});

// Clean up resources
document.Dispose();
```

---

# Spire.Doc JavaScript Caption
## Add caption to pictures in Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section
let section = document.AddSection();

// Add the first paragraph to the section
let par1 = section.AddParagraph();
par1.Format.AfterSpacing = 10; 

// Append the first picture to the paragraph
let pic1 = par1.AppendPicture({ imgFile: inputFileName });

// Set the dimensions of the first picture
pic1.Height = 100;  
pic1.Width = 120; 

// Add a caption to the first picture
let format = wasmModule.CaptionNumberingFormat.Number; 
pic1.AddCaption("Figure", format, wasmModule.CaptionPosition.BelowItem); 

// Add the second paragraph to the section
let par2 = section.AddParagraph();

// Append the second picture to the second paragraph
let pic2 = par2.AppendPicture({ imgFile: inputFileName1 });

// Set the dimensions of the second picture
pic2.Height = 100; 
pic2.Width = 120; 

// Add a caption to the second picture
pic2.AddCaption("Figure", format, wasmModule.CaptionPosition.BelowItem); 

// Update fields
document.IsUpdateFields = true;
```

---

# spire.doc javascript table caption
## add caption to table in word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Load a document from the virtual file system
document.LoadFromFile(inputFileName);

// Get the body of the first section
let body = document.Sections.get(0).Body;

// Get the first table
let table = body.Tables.get_Item(0);

// Add caption to the table
table.AddCaption("Table", wasmModule.CaptionNumberingFormat.Number, wasmModule.CaptionPosition.BelowItem);

// Update fields
document.IsUpdateFields = true;
```

---

# Spire.Doc JavaScript Cross Reference
## Create cross-reference for picture caption
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Create a new section in the document
let section = document.AddSection();

// Add the first paragraph to the section
let firstPara = section.AddParagraph();

// Add another paragraph to the section
let par1 = section.AddParagraph();
par1.Format.AfterSpacing = 10;

// Append the first picture
let pic1 = par1.AppendPicture({imgFile: inputFileName});

// Set the dimensions of the first picture
pic1.Height = 120;
pic1.Width = 120;

// Add a caption to the first picture
let format = wasmModule.CaptionNumberingFormat.Number;
let captionParagraph = pic1.AddCaption("Figure", format, wasmModule.CaptionPosition.BelowItem);

// Add an empty paragraph after the caption
section.AddParagraph();

// Add a paragraph and append the second picture
let par2 = section.AddParagraph();
let pic2 = par2.AppendPicture({imgFile: inputFileName1});

// Set the dimensions of the second picture
pic2.Height = 120;
pic2.Width = 120;

// Add a caption to the second picture
captionParagraph = pic2.AddCaption("Figure", format, wasmModule.CaptionPosition.BelowItem);

// Add an empty paragraph after the caption
section.AddParagraph();

// Create a bookmark named "Figure_2"
let bookmarkName = "Figure_2";
let paragraph = section.AddParagraph();
paragraph.AppendBookmarkStart(bookmarkName);
paragraph.AppendBookmarkEnd(bookmarkName);

// Replace the content of the bookmark
let navigator =  wasmModule.BookmarksNavigator.Create(document);
navigator.MoveToBookmark(bookmarkName);
let part = navigator.GetBookmarkContent();
part.BodyItems.Clear();
part.BodyItems.Add(captionParagraph);
navigator.ReplaceBookmarkContent({bodyPart: part});

// Create a cross-reference field pointing to the bookmark "Figure_2"
let field = wasmModule.Field.Create(document);
field.Type = wasmModule.FieldType.FieldRef;
field.Code = "REF Figure_2 \\p \\h";
firstPara.ChildObjects.Add(field);
let fieldSeparator = wasmModule.FieldMark.Create(document, wasmModule.FieldMarkType.FieldSeparator);
firstPara.ChildObjects.Add(fieldSeparator);

// Set the display text of the cross-reference field
let tr = wasmModule.TextRange.Create(document);
tr.Text = "Figure 2";
firstPara.ChildObjects.Add(tr);

let fieldEnd = wasmModule.FieldMark.Create(document, wasmModule.FieldMarkType.FieldEnd);
firstPara.ChildObjects.Add(fieldEnd);

// Update all fields in the document
document.IsUpdateFields = true;
```

---

# Document Caption with Chapter Number
## Set captions with chapter numbers for images in a Word document
```javascript
// Get the first section
let section = document.Sections.get(0);

// Define the caption name for the pictures
let name = "Caption ";

// Loop through each paragraph in the section body
for (let i = 0; i < section.Body.Paragraphs.Count; i++) {
    // Loop through each child object in the current paragraph
    for (let j = 0; j < section.Body.Paragraphs.get_Item(i).ChildObjects.Count; j++) {
        // Check if the child object is a picture
        if (section.Body.Paragraphs.get_Item(i).ChildObjects.get(j) instanceof wasmModule.DocPicture) {
            let pic1 = section.Body.Paragraphs.get_Item(i).ChildObjects.get(j); 
            let body = pic1.OwnerParagraph.Owner; 
            
            if (body != null) {
                // Get the index of the paragraph containing the picture
                let imageIndex = body.ChildObjects.IndexOf(pic1.OwnerParagraph);

                // Create a new paragraph for the caption
                let para = wasmModule.Paragraph.Create(document);
                
                // Set the caption label
                para.AppendText(name); 

                // Add a field for chapter reference
                let field1 = para.AppendField("test", wasmModule.FieldType.FieldStyleRef);
                
                // Set the code for the chapter number
                field1.Code = " STYLEREF 1 \\s "; 

                // Append a delimiter between chapter number and caption
                para.AppendText(" - "); 

                // Add a field for the picture sequence number
                let field2 = para.AppendField(name, wasmModule.FieldType.FieldSequence);
                field2.CaptionName = name; 
                field2.NumberFormat = wasmModule.CaptionNumberingFormat.Number; 

                // Insert the new caption paragraph after the picture's paragraph
                body.Paragraphs.Insert(imageIndex + 1, para);
            }
        }
    }
}

// Update all fields in the document
document.IsUpdateFields = true;
```

---

# Spire.Doc JavaScript Table Caption Cross-Reference
## Create a table with caption and cross-reference in a Word document
```javascript
// Create a new document
const document = wasmModule.Document.Create();

// Get the first section of the document
let section = document.AddSection();

// Add a new table to the section and reset its cells to 2 rows and 3 columns
let table = section.AddTable({showBorder:true});
table.ResetCells(2, 3);

// Add a caption for the table below the item
let captionParagraph = table.AddCaption("Table", wasmModule.CaptionNumberingFormat.Number, wasmModule.CaptionPosition.BelowItem);

// Define a bookmark name for the table
let bookmarkName = "Table_1";

// Add a paragraph in the section for the bookmark
let paragraph = section.AddParagraph();
paragraph.AppendBookmarkStart(bookmarkName); 
paragraph.AppendBookmarkEnd(bookmarkName);  

// Create a bookmark navigator to manage the bookmark
let navigator = wasmModule.BookmarksNavigator.Create(document);
navigator.MoveToBookmark(bookmarkName); 
let part = navigator.GetBookmarkContent(); 
part.BodyItems.Clear(); 
part.BodyItems.Add(captionParagraph);
navigator.ReplaceBookmarkContent({ bodyPart: part }); 

// Create a field for cross-referencing the bookmark
let field = wasmModule.Field.Create(document);
field.Type = wasmModule.FieldType.FieldRef; 
field.Code = "REF Table_1 \\p \\h"; 

// Add line breaks to create space
for (let i = 0; i < 3; i++) {
    paragraph.AppendBreak(wasmModule.BreakType.LineBreak); // Add line breaks
}

// Add a new paragraph for the caption cross-reference
paragraph = section.AddParagraph();
let range = paragraph.AppendText("This is a table caption cross-reference, "); 
range.CharacterFormat.FontSize = 14; 
paragraph.ChildObjects.Add(field); 

// Create a field separator for formatting
let fieldSeparator = wasmModule.FieldMark.Create(document, wasmModule.FieldMarkType.FieldSeparator);
paragraph.ChildObjects.Add(fieldSeparator); 

// Create a text range for the display text of the cross-reference
let tr = wasmModule.TextRange.Create(document);
tr.Text = "Table 1"; 
tr.CharacterFormat.FontSize = 14;
tr.CharacterFormat.TextColor = wasmModule.Color.get_DeepSkyBlue();
paragraph.ChildObjects.Add(tr); 

// Create a field end mark to close the field
let fieldEnd = wasmModule.FieldMark.Create(document, wasmModule.FieldMarkType.FieldEnd);
paragraph.ChildObjects.Add(fieldEnd); 

// Update all fields in the document
document.IsUpdateFields = true; 
```

---

# Spire.Doc JavaScript Fixed Layout
## Extract and analyze fixed layout information from Word documents
```javascript
// Create a FixedLayoutDocument from the loaded document
let layoutDoc = wasmModule.FixedLayoutDocument.Create(document);

// Get the first line in the first column of the first page
let line = layoutDoc.Pages.get_Item(0).Columns.get_Item(0).Lines.get_Item(0);

// Create a StringBuilder to store the output text
let stringBuilder = [];
stringBuilder.push("Line: " + line.Text);

// Get the paragraph that contains the line and append its text to the StringBuilder
let para = line.Paragraph;
stringBuilder.push("Paragraph text: " + para.Text + "\n");

// Get the text content of the first page
let pageText = layoutDoc.Pages.get_Item(0).Text;
stringBuilder.push(pageText + "\n");

// Iterate through each page in the FixedLayoutDocument
for (let i = 0; i < layoutDoc.Pages.Count; i++) {
    let page = layoutDoc.Pages.get_Item(i);
    // Get all the lines on the current page
    let lines = page.GetChildEntities(wasmModule.LayoutElementType.Line, true);

    // Append the page index and number of lines to the StringBuilder
    stringBuilder.push("Page " + page.PageIndex + " has " + lines.Count + " lines.\n");
}

// Append the lines of the first paragraph to the StringBuilder
// (except runs and nodes in the header and footer).
stringBuilder.push("The lines of the first paragraph:\n");
for (let i = 0; i < layoutDoc.GetLayoutEntitiesOfNode(document.FirstChild.Body.Paragraphs.get_Item(0)).Count; i++) {
    let paragraphLine = layoutDoc.GetLayoutEntitiesOfNode(document.FirstChild.Body.Paragraphs.get_Item(0)).get_Item(i);
    stringBuilder.push(paragraphLine.Text.trim() + "\n");
    const x = paragraphLine.Rectangle.X;
    const y= paragraphLine.Rectangle.Y;
    const width = paragraphLine.Rectangle.Width;
    const height = paragraphLine.Rectangle.Height;
    // Create a string with X, Y, width, and height
    const infoString = `{X: ${x}, Y: ${y}, Width: ${width}, Height: ${height}}`;
    stringBuilder.push(infoString+"\n");
}
```

---

