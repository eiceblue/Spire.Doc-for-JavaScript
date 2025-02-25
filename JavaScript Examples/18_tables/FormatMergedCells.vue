<template>
  <span>Click the following button to format the merged cells</span>
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

        //Create word document
        let document = wasmModule.Document.Create();

        //Add a new section
        let section = document.AddSection();

        //Add a new table
        let table = AddTable(section);

        //Create a new style
        let style = wasmModule.ParagraphStyle.Create(document);
        style.Name = "Style";
        style.CharacterFormat.TextColor = wasmModule.Color.get_DeepSkyBlue();
        style.CharacterFormat.Italic = true;
        style.CharacterFormat.Bold = true;
        style.CharacterFormat.FontSize = 13;
        document.Styles.Add(style);

        //Merge cell horizontally
        table.ApplyHorizontalMerge(0, 0, 1);
        //Apply style
        table.Rows.get(0).Cells.get(0).Paragraphs.get_Item(0).ApplyStyle(style.Name);
        //Set vertical and horizontal alignment
        table.Rows.get(0).Cells.get(0).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
        table.Rows.get(0).Cells.get(0).Paragraphs.get_Item(0).Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

        //Merge cell vertically
        table.ApplyVerticalMerge(0, 1, 3);
        //Apply style
        table.Rows.get(1).Cells.get(0).Paragraphs.get_Item(0).ApplyStyle(style.Name);
        //Set vertical and horizontal alignment
        table.Rows.get(1).Cells.get(0).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
        table.Rows.get(1).Cells.get(0).Paragraphs.get_Item(0).Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Left;
        //Set column width
        table.Rows.get(1).Cells.get(0).SetCellWidth(20, wasmModule.CellWidthType.Percentage);

        // Define the output file name
        const outputFileName = "FormatMergedCells_output.docx";

        // Save the document to the specified path
        document.SaveToFile({fileName: outputFileName,fileFormat: wasmModule.FileFormat.Docx2013});

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"});

        // Clean up resources
        document.Close();
        document.Dispose();

        // Download the file
        downloadName.value = outputFileName;
        downloadUrl.value = URL.createObjectURL(modifiedFile);
      }

      function AddTable(section) {
    let table = section.AddTable({showBorder: true});
    table.ResetCells(4, 3);
    //Table data
    let dt = [["Product", "", "Price"],
        ["Spire.Doc", "Pro Edition", "$799"],
        ["", "Standard Edition", "$599"],
        ["", "Free Edition", "$0"]];

    for (let r = 0; r < dt.length; r++) {
        let dataRow = table.Rows.get(r);
        dataRow.Height = 20;
        dataRow.HeightType = wasmModule.TableRowHeightType.Exactly;
        dataRow.RowFormat.BackColor = wasmModule.Color.Empty;
        for (let c = 0; c < dataRow.Cells.Count; c++) {
            if(dt[r][c] !==""){
                let range = dataRow.Cells.get(c).AddParagraph().AppendText(dt[r][c]);
                range.CharacterFormat.FontName = "Arial";
            }

        }
    }

    return table;
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
