<template>
  <span>Click the following button to set position of table in Word document as outside</span>
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
        let inputFileName = "Word.png";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`); 

        //Create a new word document and add new section
        let doc = wasmModule.Document.Create();
        let sec = doc.AddSection();

        //Get header
        let header = doc.Sections.get(0).HeadersFooters.Header;

        //Add new paragraph on header and set HorizontalAlignment of the paragraph as left
        let paragraph = header.AddParagraph();
        paragraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;

        //Load an image for the paragraph
        let headerimage = paragraph.AppendPicture({imgFile: inputFileName});
        //Add a table of 4 rows and 2 columns
        let table = header.AddTable();
        table.ResetCells(4, 2);

        //Set the position of the table to the right of the image
        table.TableFormat.WrapTextAround = true;
        table.TableFormat.Positioning.HorizPositionAbs = wasmModule.HorizontalPosition.Outside;
        table.TableFormat.Positioning.VertRelationTo = wasmModule.VerticalRelation.Margin;
        table.TableFormat.Positioning.VertPosition = 43;

        //Add contents for the table
        let data = [
            ["Spire.Doc.left", "Spire XLS.right"],
            ["Spire.Presentatio.left", "Spire.PDF.right"],
            ["Spire.DataExport.left", "Spire.PDFViewe.right"],
            ["Spire.DocViewer.left", "Spire.BarCode.right"]
        ];

        for (let r = 0; r < 4; r++) {
            let dataRow = table.Rows.get(r);
            for (let c = 0; c < 2; c++) {
                if (c == 0) {
                    let par = dataRow.Cells.get(c).AddParagraph();
                    par.AppendText(data[r][c]);
                    par.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;
                    dataRow.Cells.get(c).Width = 180;
                } else {
                    let par = dataRow.Cells.get(c).AddParagraph();
                    par.AppendText(data[r][c]);
                    par.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;
                    dataRow.Cells.get(c).Width = 180;
                }
            }
        }

        // Define the output file name
        const outputFileName = "SetOutsidePosition_output.docx";

        // Save the document to the specified path
        doc.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        doc.Close();
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
