<template>
  <span>Click the following button to create table of content in a Word document</span>
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
        let inputFileName = "Ornithogalum.jpg";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName1 = "Rosa.jpg";
        await wasmModule.FetchFileToVFS(inputFileName1,"",`${import.meta.env.BASE_URL}static/data/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName2 = "hyacinths.JPG";
        await wasmModule.FetchFileToVFS(inputFileName2,"",`${import.meta.env.BASE_URL}static/data/`);

        // Create a new document
        const document = wasmModule.Document.Create();

        // Add a new section
        let section = document.AddSection();

        // Add a new paragraph
        let para = section.AddParagraph();
        
        // Create a table of contents with default switches (\o "1-3" \h \z)
        para.AppendTOC(1, 3);
        
        // Add another paragraph for the title
        let par = section.AddParagraph();
        let tr = par.AppendText("Flowers");
        tr.CharacterFormat.FontSize = 30;
        par.Format.HorizontalAlignment = wasmModule.HorizontalAlignment.Center;

        // Create a new paragraph and set the heading level
        let para1 = section.AddParagraph();
        para1.AppendText("Ornithogalum");
        // Apply the Heading1 style to the paragraph
        para1.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading1});
        para1 = section.AddParagraph();
        // Append a picture to the paragraph
        let picture = para1.AppendPicture({imgFile: inputFileName});
        // Set the text wrapping style
        picture.TextWrappingStyle = wasmModule.TextWrappingStyle.Square;
        para1.AppendText("Ornithogalum is a genus of perennial plants mostly native to southern Europe and southern Africa belonging to the family Asparagaceae. Some species are native to other areas such as the Caucasus. Growing from a bulb, species have linear basal leaves and a slender stalk, up to 30 cm tall, bearing clusters of typically white star-shaped flowers, often striped with green.");
        para1 = section.AddParagraph();

        // Create another paragraph for the next heading
        let para2 = section.AddParagraph();
        para2.AppendText("Rosa");
        para2.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading2});
        para2 = section.AddParagraph();
        let picture2 = para2.AppendPicture({imgFile: inputFileName1});
        picture2.TextWrappingStyle = wasmModule.TextWrappingStyle.Square;
        para2.AppendText("A rose is a woody perennial flowering plant of the genus Rosa, in the family Rosaceae, or the flower it bears. There are over a hundred species and thousands of cultivars. They form a group of plants that can be erect shrubs, climbing or trailing with stems that are often armed with sharp prickles. Flowers vary in size and shape and are usually large and showy, in colours ranging from white through yellows and reds. Most species are native to Asia, with smaller numbers native to Europe, North America, and northwestern Africa. Species, cultivars and hybrids are all widely grown for their beauty and often are fragrant. Roses have acquired cultural significance in many societies. Rose plants range in size from compact, miniature roses, to climbers that can reach seven meters in height. Different species hybridize easily, and this has been used in the development of the wide range of garden roses.");
        section.AddParagraph();

        // Create another paragraph for the next heading
        let para3 = section.AddParagraph();
        para3.AppendText("Hyacinth");
        para3.ApplyStyle({builtinStyle: wasmModule.BuiltinStyle.Heading3});
        para3 = section.AddParagraph();
        let picture3 = para3.AppendPicture({imgFile: inputFileName2});
        picture3.TextWrappingStyle = wasmModule.TextWrappingStyle.Tight;
        para3.AppendText("Hyacinthus is a small genus of bulbous, fragrant flowering plants in the family Asparagaceae, subfamily Scilloideae.These are commonly called hyacinths.The genus is native to the eastern Mediterranean (from the south of Turkey through to northern Israel).");
        para3 = section.AddParagraph();
        para3.AppendText("Several species of Brodiea, Scilla, and other plants that were formerly classified in the lily family and have flower clusters borne along the stalk also have common names with the word \"hyacinth\" in them. Hyacinths should also not be confused with the genus Muscari, which are commonly known as grape hyacinths.");

        // Update TOC
        document.UpdateTableOfContents();

        // Define the output file name
        const outputFileName = "CreateTableOfContentByDefault.docx";

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
