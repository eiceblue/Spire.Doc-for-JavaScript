<template>
  <span>Click the following button to create a Cross-Reference to bookmark in Word document</span>
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

        // Create a new document
        const document = wasmModule.Document.Create();

        // Create a new section
        let section = document.AddSection();

        //Create a bookmark
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

        // Define the output file name
        const outputFileName = "CreateCrossReference.docx";

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
