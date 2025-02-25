<template>
  <span>Click the following button to create table in a Word document</span>
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

        //Open a blank Word document as template
        let document = wasmModule.Document.Create();

        let section = document.AddSection();
        addTable(section);

        // Define the output file name
        const outputFileName = "CreateTable_output.docx";

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
    };

    function addTable(section) {
    let header = ["Name", "Capital", "Continent", "Area", "Population"];
    let data =
        [
            ["Argentina", "Buenos Aires", "South America", "2777815", "32300003"],
            ["Bolivia", "La Paz", "South America", "1098575", "7300000"],
            ["Brazil", "Brasilia", "South America", "8511196", "150400000"],
            ["Canada", "Ottawa", "North America", "9976147", "26500000"],
            ["Chile", "Santiago", "South America", "756943", "13200000"],
            ["Colombia", "Bagota", "South America", "1138907", "33000000"],
            ["Cuba", "Havana", "North America", "114524", "10600000"],
            ["Ecuador", "Quito", "South America", "455502", "10600000"],
            ["El Salvador", "San Salvador", "North America", "20865", "5300000"],
            ["Guyana", "Georgetown", "South America", "214969", "800000"],
            ["Jamaica", "Kingston", "North America", "11424", "2500000"],
            ["Mexico", "Mexico City", "North America", "1967180", "88600000"],
            ["Nicaragua", "Managua", "North America", "139000", "3900000"],
            ["Paraguay", "Asuncion", "South America", "406576", "4660000"],
            ["Peru", "Lima", "South America", "1285215", "21600000"],
            ["United States of America", "Washington", "North America", "9363130", "249200000"],
            ["Uruguay", "Montevideo", "South America", "176140", "3002000"],
            ["Venezuela", "Caracas", "South America", "912047", "19700000"]
        ];
    let table = section.AddTable({showBorder: true});
    table.ResetCells(data.length + 1, header.length);

    // ***************** First Row *************************
    let row = table.Rows.get_Item(0);
    row.IsHeader = true;
    row.Height = 20;    //unit: point, 1point = 0.3528 mm
    row.HeightType = wasmModule.TableRowHeightType.Exactly;
    row.RowFormat.BackColor = wasmModule.Color.get_Gray();
    for (let i = 0; i < header.length; i++) {
        row.Cells.get(i).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
        let p = row.Cells.get(i).AddParagraph();
        p.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;
        let txtRange = p.AppendText(header[i]);
        txtRange.CharacterFormat.Bold = true;
    }

    for (let r = 0; r < data.length; r++) {
        let dataRow = table.Rows.get(r + 1);
        dataRow.Height = 20;
        dataRow.HeightType = wasmModule.TableRowHeightType.Exactly;
        dataRow.RowFormat.BackColor = wasmModule.Color.Empty();
        for (let c = 0; c < data[r].length; c++) {
            dataRow.Cells.get(c).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
            dataRow.Cells.get(c).AddParagraph().AppendText(data[r][c]);
        }
    }

    for (let j = 1; j < table.Rows.Count; j++) {
        if (j % 2 == 0) {
            let row2 = table.Rows.get_Item(j);
            for (let f = 0; f < row2.Cells.Count; f++) {
                row2.Cells.get(f).CellFormat.BackColor = wasmModule.Color.get_LightBlue();
            }
        }
    }
}


    return {
      startProcessing,
      downloadName,
      downloadUrl,
    };
  },
};
</script>
