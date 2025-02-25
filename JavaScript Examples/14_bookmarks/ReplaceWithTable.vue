<template>
  <span>The example shows how to replace bookmark with a table in the Word document.</span>
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
        let inputFileName = "ReplaceWithTable.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const doc = wasmModule.Document.Create();

        // Load a document from the virtual file system
        doc.LoadFromFile(inputFileName);

        //Create a table
        let table = wasmModule.Table.Create(doc, true);
        table.ResetCells(4,5);
        //Create datatable
        let dt = [       //not supported datatable
            //["id", "name", "job", "email", "salary"],
            ["Name", "Capital", "Continent", "Area", "Population"],
            ["Argentina", "Buenos Aires", "South America", "2777815", "32300003"],
            ["Bolivia", "La Paz", "South America", "1098575", "7300000"],
            ["Brazil", "Brasilia", "South America", "8511196", "150400000"]];
        //Fill the table with the data of datatable
        for (let i = 0; i < 4; i++) {
            for (let j = 0; j < 5; j++) {
                table.Rows.get(i).Cells.get(j).AddParagraph().AppendText(dt[i][j]);
              }
        }

        //Get the specific bookmark by its name
        let navigator = wasmModule.BookmarksNavigator.Create(doc);
        navigator.MoveToBookmark("Test");

        //Create a TextBodyPart instance and add the table to it
        let part = wasmModule.TextBodyPart.Create(doc);
        part.BodyItems.Add(table);

        //Replace the current bookmark content with the TextBodyPart object
        navigator.ReplaceBookmarkContent({bodyPart : part});
        
        // Define the output file name
        const outputFileName = "ReplaceWithTable-result.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Clean up resources
        doc.Dispose();

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
