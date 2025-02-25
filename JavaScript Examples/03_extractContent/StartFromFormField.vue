<template>
  <span>The following example demonstrates how to extract content start from a form field. </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"> Click here to download the generated file </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref('');

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = 'TextInputField.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        //Create the source document
        let sourceDocument = wasmModule.Document.Create();

        //Load the source document from disk.
        sourceDocument.LoadFromFile(inputFileName);

        //Create a destination document
        let destinationDoc = wasmModule.Document.Create();

        //Add a section
        let section = destinationDoc.AddSection();

        //Define a variables
        let index = 0;
        let formFields = sourceDocument.Sections.get(0).Body.FormFields;
        //Traverse FormFields
        for (let i = 0; i < formFields.Count; i++) {
          let field = formFields.get_Item(i);
          //Find FieldFormTextInput type field
          if (field.Type == wasmModule.FieldType.FieldFormTextInput) {
            //Get the paragraph
            let paragraph = field.OwnerParagraph;

            //Get the index
            index = sourceDocument.Sections.get(0).Body.ChildObjects.IndexOf(paragraph);
            break;
          }
        }

        //Extract the content
        for (let i = index; i < index + 3; i++) {
          //Clone the ChildObjects of source document
          let doobj = sourceDocument.Sections.get(0).Body.ChildObjects.get(i).Clone();

          //Add to destination document
          section.Body.ChildObjects.Add(doobj);
        }

        // Define the output file name
        const outputFileName = 'StartFromFormField.docx';

        // Save the document to the specified path
        destinationDoc.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        // Clean up resources
        destinationDoc.Dispose();

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
