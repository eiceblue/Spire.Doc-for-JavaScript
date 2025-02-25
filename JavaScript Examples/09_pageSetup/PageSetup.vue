<template>
  <span>Click the following button to set Word document page properties.</span>
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

        function addTable(section) {
          let header = ["Name", "Capital", "Continent", "Area", "Population"];
          let data =
            [
              ["Argentina", "Buenos Aires", "South America", "2777815", "32300003"],
              ["Bolivia", "La Paz", "South", "1098575", "7300000"],
              ["Brazil", "Brasilia", "South", "8511196", "150400000"],
              ["Canada", "Ottawa", "North", "9976147", "26500000"],
              ["Chile", "Santiago", "South", "756943", "13200000"],
              ["Colombia", "Bagota", "South", "1138907", "33000000"],
              ["Cuba", "Havana", "North", "114524", "10600000"],
              ["Ecuador", "Quito", "South", "455502", "10600000"],
              ["El Salvador", "San Salvador", "North", "20865", "5300000"],
              ["Guyana", "Georgetown", "South", "214969", "800000"],
              ["Jamaica", "Kingston", "North", "11424", "2500000"],
              ["Mexico", "Mexico City", "North", "1967180", "88600000"],
              ["Nicaragua", "Managua", "North", "139000", "3900000"],
              ["Paraguay", "Asuncion", "South", "406576", "4660000"],
              ["Peru", "Lima", "South", "1285215", "21600000"],
              ["United States", "Washington", "North", "9363130", "249200000"],
              ["Uruguay", "Montevideo", "South", "176140", "3002000"],
              ["Venezuela", "Caracas", "South", "912047", "19700000"]
            ];
          let table = section.AddTable({showBorder:true});
          table.ResetCells(data.length + 1, header.length);

          // ***************** First Row *************************
          let row = table.Rows.get(0);
          row.IsHeader = true;
          row.Height = 20;
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
            dataRow.RowFormat.BackColor = wasmModule.Color.Empty;
            for (let c = 0; c < data[r].length; c++) {
              dataRow.Cells.get(c).CellFormat.VerticalAlignment = wasmModule.VerticalAlignment.Middle;
              dataRow.Cells.get(c).AddParagraph().AppendText(data[r][c]);
            }
          }
        }

        function InsertHeaderAndFooter(section,headerPic, footerPic) {
          let header = section.HeadersFooters.Header;
          let footer = section.HeadersFooters.Footer;

          //Insert picture and text to header.
          let headerParagraph = header.AddParagraph();


          let headerPicture = headerParagraph.AppendPicture({ imgFile: headerPic });

          //Header text.
          let text = headerParagraph.AppendText("Demo of Spire.Doc");
          text.CharacterFormat.FontName = "Arial";
          text.CharacterFormat.FontSize = 10;
          text.CharacterFormat.Italic = true;
          headerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

          //Border.
          headerParagraph.Format.Borders.Bottom.BorderType = wasmModule.BorderStyle.Single;
          headerParagraph.Format.Borders.Bottom.Space = 0.05;


          //Header picture layout - text wrapping.
          headerPicture.TextWrappingStyle = wasmModule.TextWrappingStyle.Behind;

          //Header picture layout - position.
          headerPicture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
          headerPicture.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
          headerPicture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
          headerPicture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Top;

          //Insert picture to footer.
          let footerParagraph = footer.AddParagraph();
          let footerPicture = footerParagraph.AppendPicture({ imgFile: footerPic });

          //Footer picture layout.
          footerPicture.TextWrappingStyle = wasmModule.TextWrappingStyle.Behind;
          footerPicture.HorizontalOrigin = wasmModule.HorizontalOrigin.Page;
          footerPicture.HorizontalAlignment = wasmModule.ShapeHorizontalAlignment.Left;
          footerPicture.VerticalOrigin = wasmModule.VerticalOrigin.Page;
          footerPicture.VerticalAlignment = wasmModule.ShapeVerticalAlignment.Bottom;

          //Insert page number.
          footerParagraph.AppendField("page number", wasmModule.FieldType.FieldPage);
          footerParagraph.AppendText(" of ");
          footerParagraph.AppendField("number of pages", wasmModule.FieldType.FieldNumPages);
          footerParagraph.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Right;

          //Border.
          footerParagraph.Format.Borders.Top.BorderType = wasmModule.BorderStyle.Single;
          footerParagraph.Format.Borders.Top.Space = 0.05;

        }


        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF", "/Library/Fonts/", `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName1 = "Header.png";
        await wasmModule.FetchFileToVFS(inputFileName1, "", `${import.meta.env.BASE_URL}static/data/`);
        let inputFileName2 = "Footer.png";
        await wasmModule.FetchFileToVFS(inputFileName2, "", `${import.meta.env.BASE_URL}static/data/`);

        //Create Word document.
        let document = wasmModule.Document.Create();
        let section = document.AddSection();

        //The unit of all measures below is point, 1point = 0.3528 mm.
        section.PageSetup.PageSize = wasmModule.PageSize.A4();
        section.PageSetup.Margins.Top = 72;
        section.PageSetup.Margins.Bottom = 72;
        section.PageSetup.Margins.Left = 89.85;
        section.PageSetup.Margins.Right = 89.85;

        //Insert header and footer.
        InsertHeaderAndFooter(section,inputFileName1,inputFileName2);

        addTable(section);

        // Define the output file name
        const outputFileName = "PageSetup_out.docx";

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
