<template>
  <span>The following example shows how to replace text with filed.</span>
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
        let inputFileName = 'ReplaceTextWithField.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

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

        // Define the output file name
        const outputFileName = 'ReplaceTextWithField.docx';

        // Save the document to the specified path
        doc.SaveToFile({
          fileName: outputFileName,
          fileFormat: wasmModule.FileFormat.Docx2013,
        });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        });

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
