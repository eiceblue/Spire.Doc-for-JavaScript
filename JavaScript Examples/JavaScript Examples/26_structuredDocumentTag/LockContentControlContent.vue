<template>
  <span>Click the following button to lock the content of content controls in a Word document</span>
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

        // Create a new document
        const document = wasmModule.Document.Create();

        let htmlString = "<table style=\"width: 100 % \">"
        + "<tr><th> Number </th><th> Name </th ><th>Age</th ></tr>"
        + "<tr><td> 1 </td><td> Smith </td><td> 50 </td></tr>"
        + "<tr> <td> 2 </td><td> Jackson </td><td> 94 </td> </tr>"
        + "</table>";
        
        let section = document.AddSection();
        let paragraph = section.AddParagraph();
        paragraph.AppendHTML(htmlString);

        //Create StructureDocumentTag for document
        let sdt = wasmModule.StructureDocumentTag.Create(document);
        let section2 = document.AddSection();
        section2.Body.ChildObjects.Add(sdt);

        //Specify the type
        sdt.SDTProperties.SDTType = wasmModule.SdtType.RichText;

        for (let i = 0; i < section.Body.ChildObjects.Count; i++) {
            let obj = section.Body.ChildObjects.get(i);
            if (obj.DocumentObjectType === wasmModule.DocumentObjectType.Table) {
                sdt.SDTContent.ChildObjects.Add(obj.Clone());
            }
        }

        // Lock content
        sdt.SDTProperties.LockSettings = wasmModule.LockSettingsType.ContentLocked;

        document.Sections.Remove(section);
        
        // Define the output file name
        const outputFileName = "LockContentControlContent.docx";

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
