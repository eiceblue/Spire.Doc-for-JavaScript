<template>
  <span>Click the following button to extract convert field to body text in a Word document</span>
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

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "TextInputField.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        // Load a document from the virtual file system
        doc.LoadFromFile(inputFileName);

        // Traverse FormFields
        for (let i = 0;i < doc.Sections.get(0).Body.FormFields.Count;i++) {
          let field =doc.Sections.get(0).Body.FormFields.get_Item(i);

          // Find FieldFormTextInput type field
          if (field.Type == wasmModule.FieldType.FieldFormTextInput) {
            // Get the paragraph
            let paragraph = field.OwnerParagraph;

            // Define variables
            let startIndex = 0;
            let endIndex = 0;

            // Create a new TextRange
            let textRange = wasmModule.TextRange.Create(doc);

            //Set text for textRange
            textRange.Text = paragraph.Text;

            // Traverse DocumentObjectS of field paragraph
            for (let j = 0; j < paragraph.ChildObjects.Count; j++) {
              let obj = paragraph.ChildObjects.get(j);
              // If its DocumentObjectType is BookmarkStart
              if (obj.DocumentObjectType ===wasmModule.DocumentObjectType.BookmarkStart) {
                // Get the index
                startIndex = paragraph.ChildObjects.IndexOf(obj);
              }
              // If its DocumentObjectType is BookmarkEnd
              if (obj.DocumentObjectType ===wasmModule.DocumentObjectType.BookmarkEnd) {
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

        // Define the output file name
        const outputFileName = "ConvertFieldToBodyText.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        doc.Dispose();

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
