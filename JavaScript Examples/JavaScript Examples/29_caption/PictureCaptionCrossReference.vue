<template>
  <span>Click the following button to create cross-reference for picture caption</span>
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
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "Spire.Doc.png";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        
        let inputFileName1 = "Word.png";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);
        
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
        
        // Define the output file name
        const outputFileName = "PictureCaptionCrossReference.docx";

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

    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
