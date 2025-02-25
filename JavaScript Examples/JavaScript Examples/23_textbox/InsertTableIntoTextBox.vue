<template>
  <span>Click the following button to insert table into text box in a Word document</span>
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
        let inputFileName = "Template.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Create a new section
        let section = document.AddSection();

        // Add a paragraph to the section
        let paragraph = section.AddParagraph();

        // Add a textbox to the paragraph
        let textbox = paragraph.AppendTextBox(300, 100);

        // Set the position of the textbox
        textbox.Format.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
        textbox.Format.HorizontalPosition = 140;
        textbox.Format.VerticalOrigin = wasmModule.VerticalOrigin.Page;
        textbox.Format.VerticalPosition = 50;

        // Add text to the textbox
        let textboxParagraph = textbox.Body.AddParagraph();
        let textboxRange = textboxParagraph.AppendText("Table 1");
        textboxRange.CharacterFormat.FontName = "Arial";

        // Insert table to the textbox
        let table = textbox.Body.AddTable({showBorder: true});

        // Specify the number of rows and columns of the table
        table.ResetCells(4, 4);

        let data =
            [
                ["Name", "Age", "Gender", "ID"],
                ["John", "28", "Male", "0023"],
                ["Steve", "30", "Male", "0024"],
                ["Lucy", "26", "female", "0025"]
            ];

        // Add data to the table
        for (let i = 0; i < 4; i++) {
            for (let j = 0; j < 4; j++) {
                let tableRange = table.Rows.get(i).Cells.get(j).AddParagraph().AppendText(data[i][j]);
                tableRange.CharacterFormat.FontName = "Arial";
            }
        }

        // Apply style to the table
        table.ApplyStyle(wasmModule.DefaultTableStyle.TableColorful2);

        // Define the output file name
        const outputFileName = "InsertTableIntoTextBox.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

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
