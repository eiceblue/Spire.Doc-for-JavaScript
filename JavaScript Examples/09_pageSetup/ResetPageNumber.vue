<template>
  <span>Click the following button to reset page number for each section start at 1.</span>
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
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName1 = "ResetPageNumber1.docx";
        await wasmModule.FetchFileToVFS(inputFileName1, "", `${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2 = "ResetPageNumber2.docx";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);
        let inputFileName3 = "ResetPageNumber3.docx";
        await wasmModule.FetchFileToVFS(inputFileName3, "", `${import.meta.env.BASE_URL}static/data/`);


        //Create three Word documents and load three different Word documents 
        let document1 = wasmModule.Document.Create();
        document1.LoadFromFile(inputFileName1);

        let document2 = wasmModule.Document.Create();
        document2.LoadFromFile(inputFileName2);

        let document3 = wasmModule.Document.Create();
        document3.LoadFromFile(inputFileName3);

        //Use section method to combine all documents into one word document.
        for (let i = 0; i < document2.Sections.Count; i++) {
          let sec = document2.Sections.get(i);
          document1.Sections.Add(sec.Clone());
        }
        for (let i = 0; i < document3.Sections.Count; i++) {
          let sec = document3.Sections.get(i);
          document1.Sections.Add(sec.Clone());
        }

        //Traverse every section of document1.
        for (let i = 0; i < document1.Sections.Count; i++) {
          let sec = document1.Sections.get_Item(i);
          //Traverse every object of the footer.
          for (let j = 0; j < sec.HeadersFooters.Footer.ChildObjects.Count; j++) {
            let obj = sec.HeadersFooters.Footer.ChildObjects.get(j);
            if (obj.DocumentObjectType == wasmModule.DocumentObjectType.StructureDocumentTag) {
              let para = obj.ChildObjects.get(0);
              for (let k = 0; k < para.ChildObjects.Count; k++) {
                let item = para.ChildObjects.get(k);
                if (item.DocumentObjectType == wasmModule.DocumentObjectType.Field)
                  //Find the item and its field type is FieldNumPages.
                  if (item.Type == wasmModule.FieldType.FieldNumPages) {
                    //Change field type to FieldSectionPages.
                    item.Type = FieldType.FieldSectionPages;
                  }
              }
            }
          }
        }

        //Restart page number of section and set the starting page number to 1.
        document1.Sections.get(1).PageSetup.RestartPageNumbering = true;
        document1.Sections.get(1).PageSetup.PageStartingNumber = 1;

        document1.Sections.get(2).PageSetup.RestartPageNumbering = true;
        document1.Sections.get(2).PageSetup.PageStartingNumber = 1;

        // Define the output file name
        const outputFileName = "ResetPageNumber_out.docx";

        //Save to file.
        document1.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

        // Clean up resources
        document1.Dispose();
        document2.Dispose();
        document3.Dispose();

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
