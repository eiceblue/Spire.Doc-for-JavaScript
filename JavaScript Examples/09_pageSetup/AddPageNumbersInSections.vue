<template>
  <span>Click the following button to add page numbers in different sections in Word document.</span>
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
        let inputFileName = "Template_Docx_4.docx";
        await wasmModule.FetchFileToVFS(inputFileName, "", `${import.meta.env.BASE_URL}static/data/`);

        //Create Word document.
        let document = wasmModule.Document.Create();

        //Load the file 
        document.LoadFromFile(inputFileName);

        //Repeat step2 and Step3 for the rest sections, so change the code with for loop.
        for (let i = 0; i < 3; i++) {
          let footer = document.Sections.get(i).HeadersFooters.Footer;
          let footerParagraph = footer.AddParagraph();
          footerParagraph.AppendField("page number", wasmModule.FieldType.FieldPage);
          footerParagraph.AppendText(" of ");
          footerParagraph.AppendField("number of pages", wasmModule.FieldType.FieldSectionPages);
          footerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

          if (i == 2)
            break;
          else {
            document.Sections.get(i + 1).PageSetup.RestartPageNumbering = true;
            document.Sections.get(i + 1).PageSetup.PageStartingNumber = 1;
          }
        }

        // Define the output file name
        const outputFileName = "AddPageNumbersInSections_out.docx";

        // Save the document to the specified path
        document.SaveToFile({ fileName: outputFileName, fileFormat: wasmModule.FileFormat.Docx2013 });

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });

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
