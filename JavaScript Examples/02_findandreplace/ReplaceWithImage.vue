<template>
  <span>The following example shows how to replace text with image in a Word document. </span>
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
        let pngName = 'E-iceblue.png';
        await wasmModule.FetchFileToVFS(pngName, '', `${import.meta.env.BASE_URL}static/data/`);

        let inputFileName = 'Template.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);
        //Load a template document
        let document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

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

        // Define the output file name
        const outputFileName = 'ReplaceWithImage.docx';

        // Save the document to the specified path
        document.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

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
