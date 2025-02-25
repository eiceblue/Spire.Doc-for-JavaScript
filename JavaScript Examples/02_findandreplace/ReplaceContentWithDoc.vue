<template>
  <span>The following example demonstrates how to replace content with a document.</span>
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

        // Load the sample files into the virtual file system (VFS)
        let inputFileName1 = 'ReplaceContentWithDoc.docx';
        await wasmModule.FetchFileToVFS(inputFileName1, '', `${import.meta.env.BASE_URL}static/data/`);

        let inputFileName2 = 'Insert.docx';
        await wasmModule.FetchFileToVFS(inputFileName2, '', `${import.meta.env.BASE_URL}static/data/`);
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

        // Define the output file name
        const outputFileName = 'ReplaceContentWithDoc.docx';

        // Save the document to the specified path
        document1.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

        // Clean up resources
        document1.Dispose();

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
