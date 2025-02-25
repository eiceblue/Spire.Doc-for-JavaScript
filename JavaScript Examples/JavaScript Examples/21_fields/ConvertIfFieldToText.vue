<template>
  <span>Click the following button to convert if field to text in a Word document</span>
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
        let inputFileName = "IfFieldSample.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

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

        // Define the output file name
        const outputFileName = "ConvertIfFieldToText.docx";

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
