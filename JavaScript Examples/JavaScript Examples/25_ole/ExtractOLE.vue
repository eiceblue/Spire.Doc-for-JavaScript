<template>
  <span>Click the following button to extract OLE from a Word document</span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="PdfDownloadUrl" :href="PdfDownloadUrl" :download="PdfDownloadName">
    Click here to download the extracted PDF file
  </a>
  <a v-if="XlsDownloadUrl" :href="XlsDownloadUrl" :download="XlsDownloadName">
    Click here to download the extracted XLS file
  </a>
  <a v-if="PPTDownloadUrl" :href="PPTDownloadUrl" :download="PPTDownloadName">
    Click here to download the extracted PPT file
  </a>
</template>

<script>
import { ref } from "vue";

export default {
  setup() {
    const PdfDownloadUrl = ref(null);
    const PdfDownloadName = ref("");
    const XlsDownloadUrl = ref(null);
    const XlsDownloadName = ref("");
    const PPTDownloadUrl = ref(null);
    const PPTDownloadName = ref("");

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS("ARIALUNI.TTF","/Library/Fonts/",`${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = "OLEs.docx";
        await wasmModule.FetchFileToVFS(inputFileName,"",`${import.meta.env.BASE_URL}static/data/`);

        // Define the output file name
        const PdfOutputFileName = "ExtractOLE.pdf";
        const XlsOutputFileName = "ExtractOLE.xls";
        const PPTOutputFileName = "ExtractOLE.pptx";

        // Create a new document
        const document = wasmModule.Document.Create();

        // Load a document from the virtual file system
        document.LoadFromFile(inputFileName);

        // Traverse through all sections of the word document
        for (let i = 0; i < document.Sections.Count; i++) {
            let sec = document.Sections.get(i);
            // Traverse through all Child Objects in the body of each section
            for (let j = 0; j < sec.Body.ChildObjects.Count; j++) {
                let obj = sec.Body.ChildObjects.get(j);
                // Find the paragraph
                if (obj instanceof wasmModule.Paragraph) {
                    let par = obj;
                    for (let k = 0; k < par.ChildObjects.Count; k++) {
                        let o = par.ChildObjects.get(k);
                        // Check whether the object is OLE
                        if (o.DocumentObjectType === wasmModule.DocumentObjectType.OleObject) {
                            let Ole = o;
                            let s = Ole.ObjectType;

                            // Check whether the object type is "Acrobat.Document.11"
                            if (s === "Acrobat.Document.DC") {
                                // Write the data of OLE into file
                                wasmModule.FS.writeFile(PdfOutputFileName, Ole.NativeData);
                            }

                            // Check whether the object type is "Excel.Sheet.8"
                            else if (s === "Excel.Sheet.8") {
                                // Write the data of OLE into file
                                wasmModule.FS.writeFile(XlsOutputFileName, Ole.NativeData);
                            }

                            //  Check whether the object type is "PowerPoint.Show.12"
                            else if (s === "PowerPoint.Show.12") {
                                // Write the data of OLE into file
                                wasmModule.FS.writeFile(PPTOutputFileName, Ole.NativeData);
                            }
                          }
                      }
                  }
              }
          }

        // Clean up resources
        document.Dispose();
        
        // Read the saved file and convert to a Blob object
        const modifiedFileArrayPdf = wasmModule.FS.readFile(PdfOutputFileName);
        const modifiedFilePdf = new Blob([modifiedFileArrayPdf], {type: "application/pdf"});
        
        // Read the saved file and convert to a Blob object
        const modifiedFileArrayXls = wasmModule.FS.readFile(XlsOutputFileName);
        const modifiedFileXls  = new Blob([modifiedFileArrayXls], {type: "application/vnd.ms-excel"});

        // Read the saved file and convert to a Blob object
        const modifiedFileArrayPPT = wasmModule.FS.readFile(PPTOutputFileName);
        const modifiedFilePPT = new Blob([modifiedFileArrayPPT], {type: "application/vnd.openxmlformats-officedocument.presentationml.presentation"});

        // Download the file
        PdfDownloadName.value = PdfOutputFileName;
        PdfDownloadUrl.value = URL.createObjectURL(modifiedFilePdf);

        // Download the file
        XlsDownloadName.value = XlsOutputFileName;
        XlsDownloadUrl.value = URL.createObjectURL(modifiedFileXls);

        // Download the file
        PPTDownloadName.value = PPTOutputFileName;
        PPTDownloadUrl.value = URL.createObjectURL(modifiedFilePPT);
      }
    };

    return {
      startProcessing,
      PdfDownloadName,
      PdfDownloadUrl,
      XlsDownloadUrl,
      XlsDownloadName,
      PPTDownloadUrl,
      PPTDownloadName
    };
  },
};
</script>
