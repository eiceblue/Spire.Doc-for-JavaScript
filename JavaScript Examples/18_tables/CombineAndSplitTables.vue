<template>
  <span>Click the following button to combine and split tables in a Word document</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName">
    Click here to download the generated file
  </a>
</template>

<script>
import { ref } from "vue";
import JSZip from "jszip";

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
        let inputFileName = "CombineAndSplitTables.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);         

        let outputDirectoryName = "outputFiles/";
        FS.mkdirTree(outputDirectoryName);

        //Load document from disk
        let doc = wasmModule.Document.Create();
        doc.LoadFromFile(inputFileName);

        //Get the first section
        let section = doc.Sections.get_Item(0);

        //Get the first and second table
        let table1 = section.Tables.get_Item(0);
        let table2 = section.Tables.get_Item(1);

        //Add the rows of table2 to table1
        for (let i = 0; i < table2.Rows.Count; i++) {
            table1.Rows.Add(table2.Rows.get(i).Clone());
        }

        //Remove the table2
        section.Tables.Remove(table2);

        // Define the output file name
        const outputFileName = "CombineTables_output.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputDirectoryName +outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});


        SplitTable(inputFileName)

        const zip = new JSZip();
        let items = await FS.readdir(outputDirectoryName);
        items = items.filter((item) => item !== "." && item !== "..");
        for (const item of items) {
          const itemPath = `${outputDirectoryName}/${item}`;
          const fileData = await FS.readFile(itemPath);
          zip.file(item, fileData);
        }

        // Zip files 
        const zipBlob = await zip.generateAsync({ type: "blob" });
        const zipDownloadUrl = URL.createObjectURL(zipBlob);
        const zipDownloadName = `outputFiles.zip`;


        // Clean up resources
        doc.Close();
        doc.Dispose();

        // download zip
        downloadName.value = zipDownloadName;
        downloadUrl.value = zipDownloadUrl;
      }

      function SplitTable(inputFileName) {

    //Load document from disk
    let doc = wasmModule.Document.Create();
    doc.LoadFromFile(inputFileName);

    //Get the first section
    let section = doc.Sections.get(0);

    //Get the first table
    let table = section.Tables.get_Item(0);

    //We will split the table at the third row;
    let splitIndex = 2;

    //Create a new table for the split table
    let newTable = wasmModule.Table.Create(section.Document, false);

    //Add rows to the new table
    for (let i = splitIndex; i < table.Rows.Count; i++) {
        newTable.Rows.Add(table.Rows.get(i).Clone());
    }

    //Remove rows from original table
    for (let i = table.Rows.Count - 1; i >= splitIndex; i--) {
        table.Rows.RemoveAt(i);
    }

    //Add the new table in section
    section.Tables.Add(newTable);

    const outFileName ="SplitTables_output.docx"
    //Save the Word file
    doc.SaveToFile({fileName: "outputFiles/" +outFileName, fileFormat: wasmModule.FileFormat.Docx2013});
    
    doc.Close();
    doc.Dispose();
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
