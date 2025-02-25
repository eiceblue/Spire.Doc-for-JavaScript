<template>
  <span>The following example demonstrates how to extract content between paragraph styles. </span>
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
        let inputFileName = 'BetweenParagraphStyle.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const sourceDocument = wasmModule.Document.Create();

        // Load a document from the virtual file system
        sourceDocument.LoadFromFile(inputFileName);

        //Create a destination document
        let destinationDoc = wasmModule.Document.Create();

        //Add a section
        let section = destinationDoc.AddSection();
        let stylename1 = '1';
        let stylename2 = '2';
        //Extract content between the first paragraph to the third paragraph
        let startindex = 0;
        let endindex = 0;
        //travel the sections of source document
        for (let i = 0; i < sourceDocument.Sections.Count; i++) {
          let section1 = sourceDocument.Sections.get(i);
          //travel the paragraphs
          for (let j = 0; j < section1.Paragraphs.Count; j++) {
            let paragraph = section1.Paragraphs.get_Item(j);
            //Judge paragraph style1
            if (paragraph.StyleName === stylename1) {
              //Get the paragraph index
              startindex = section1.Body.Paragraphs.IndexOf(paragraph);
            }
            //Judge paragraph style2
            if (paragraph.StyleName === stylename2) {
              //Get the paragraph index
              endindex = section1.Body.Paragraphs.IndexOf(paragraph);
            }
          }
          //Extract the content
          for (let i = startindex + 1; i < endindex; i++) {
            //Clone the ChildObjects of source document
            let doobj = sourceDocument.Sections.get_Item(0).Body.ChildObjects.get_Item(i).Clone();

            //Add to destination document
            destinationDoc.Sections.get_Item(0).Body.ChildObjects.Add(doobj);
          }
        }

        // Define the output file name
        const outputFileName = 'BetweenParagraphStyles.docx';

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
