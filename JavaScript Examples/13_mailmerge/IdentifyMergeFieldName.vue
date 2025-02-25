<template>
  <span>The example shows how to identify merge field names in Word document.</span>
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
        let inputFileName = "IdentifyMergeFieldName.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Load the file from disk.
        document.LoadFromFile(inputFileName);

        //Get the collection of group names.
        let GroupNames = document.MailMerge.GetMergeGroupNames();

        //Get the collection of merge field names in a specific group.
        let MergeFieldNamesWithinRegion = document.MailMerge.GetMergeFieldNames({groupName: "Products"});

        //Get the collection of all the merge field names.
        let MergeFieldNames = document.MailMerge.GetMergeFieldNames();

        let content = [];
        content.push("----------------Group Names-----------------------------------------\n");
        for (let i = 0; i < GroupNames.length; i++) {
            content.push(GroupNames[i] + "\n");
        }

        content.push("----------------Merge field names within a specific group-----------\n");
        for (let j = 0; j < MergeFieldNamesWithinRegion.length; j++) {
            content.push(MergeFieldNamesWithinRegion[j] + "\n");
        }

        content.push("----------------All of the merge field names------------------------\n");
        for (let k = 0; k < MergeFieldNames.length; k++) {
            content.push(MergeFieldNames[k] + "\n");
        }

        // Define the output file name
        const outputFileName = "IdentifyMergeFieldName-result.txt";

        //Save to file.
        wasmModule.FS.writeFile(outputFileName, content.join("\n"));
        document.Close();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "text/plain"});

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
