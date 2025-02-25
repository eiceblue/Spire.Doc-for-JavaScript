<template>
  <span>Click the following button to create form field</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);
        
        // Load the XML file into the virtual file system (VFS)
        const inputFileName = "Form.xml";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Load the images into the virtual file system (VFS)
        let headerImage = "Header.png";
        await wasmModule.FetchFileToVFS(headerImage,"",`${import.meta.env.BASE_URL}static/data/`);
        
        const footerImage = "Footer.png";
        await wasmModule.FetchFileToVFS(footerImage,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Create a new section
        let section = document.AddSection();
        
        // Insert header and footer
        InsertHeaderAndFooter(section,headerImage,footerImage);

        // Add title
        AddTitle(section);

        // Add form
        AddForm(section,inputFileName);

        // Define the output file name
        const outputFileName = "CreateFormField.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        document.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }
    };
    
    // Function to insert header and footer into a section
    const InsertHeaderAndFooter = (section,inputFileName,inputFileName_1) => {
      // Insert picture and text into the header
      let headerParagraph = section.HeadersFooters.Header.AddParagraph();

      // Append an image to the header paragraph
      let headerPicture = headerParagraph.AppendPicture({imgFile: inputFileName});

      // Append text to the header paragraph
      let text = headerParagraph.AppendText("Demo of Spire.Doc");
      text.CharacterFormat.FontName = "Arial";
      text.CharacterFormat.FontSize = 10;
      text.CharacterFormat.Italic = true;
      headerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

      // Set bottom border for the header paragraph
      headerParagraph.Format.Borders.Bottom.BorderType = wasmModule.BorderStyle.Single;
      headerParagraph.Format.Borders.Bottom.Space = 0.05;

      // Configure the picture's text wrapping style
      headerPicture.TextWrappingStyle = wasmModule.TextWrappingStyle.Behind;

      // Set the position of the header picture
      headerPicture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
      headerPicture.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
      headerPicture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
      headerPicture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Top;

      // Insert picture and text into the footer
      let footerParagraph = section.HeadersFooters.Footer.AddParagraph();
      let footerPicture = footerParagraph.AppendPicture({imgFile: inputFileName_1});

      // Configure the picture's text wrapping style in the footer
      footerPicture.TextWrappingStyle = wasmModule.TextWrappingStyle.Behind;

      // Set the position of the footer picture
      footerPicture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
      footerPicture.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
      footerPicture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
      footerPicture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Bottom;

      // Add page number field to the footer paragraph
      footerParagraph.AppendField("page number", wasmModule.FieldType.FieldPage);
      footerParagraph.AppendText(" of ");
      footerParagraph.AppendField("number of pages", wasmModule.FieldType.FieldNumPages);
      footerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

      footerParagraph.Format.Borders.Top.BorderType = wasmModule.BorderStyle.Single;
      footerParagraph.Format.Borders.Top.Space = 0.05;
    };

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
    const AddForm = (section,inputFileName) => {
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

      // Read XML content from the input file
      const data = wasmModule.FS.readFile(inputFileName);
        
      // Create a TextDecoder instance to decode the Uint8Array into a string using UTF-8 encoding
      const decoder = new TextDecoder('utf-8');
        
      // Decode the Uint8Array data into a string
      const stringData = decoder.decode(data);

      // Parse the decoded string as XML using DOMParser
      const xmlDocument = new DOMParser().parseFromString(stringData,'application/xml');

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

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
