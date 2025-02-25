<template>
  <span>The below example shows how to get the revisions details of the paragraph in the document. </span>
  <el-button @click="startProcessing">Start</el-button>
  <a v-if="downloadUrl" :href="downloadUrl" :download="downloadName"> Click here to download the generated file </a>
</template>

<script>
import { ref } from 'vue';

export default {
  setup() {
    const downloadUrl = ref(null);
    const downloadName = ref('');

    const startProcessing = async () => {
      wasmModule = window.wasmModule;
      if (wasmModule) {
        // Load the ARIALUNI.TTF font file into the virtual file system (VFS)
        await wasmModule.FetchFileToVFS('ARIALUNI.TTF', '/Library/Fonts/', `${import.meta.env.BASE_URL}static/font/`);

        // Load the sample file into the virtual file system (VFS)
        let inputFileName = 'Revisions.docx';
        await wasmModule.FetchFileToVFS(inputFileName, '', `${import.meta.env.BASE_URL}static/data/`);

        let document = wasmModule.Document.Create();
        document.LoadFromFile(inputFileName);

        let builder = [];

        //loop paragraph
        for (let i = 0; i < document.Sections.Count; i++) {
          let section = document.Sections.get_Item(0);
          for (let j = 0; j < section.Paragraphs.Count; j++) {
            let paragraph = section.Paragraphs.get_Item(j);
            // Check if the Paragraph is a deleted revision.
            if (paragraph.IsDeleteRevision) {
              // Append information about the deleted revision to the builder.
              builder.push('The section' + document.GetIndex(section) + 'paragraph' + section.GetIndex(paragraph) + 'has been changed (deleted).' + '\n');
              builder.push('Author: ' + paragraph.DeleteRevision.Author + '\n');
              builder.push('DateTime: ' + paragraph.DeleteRevision.DateTime.ToString() + '\n');
              builder.push('Type: ' + paragraph.DeleteRevision.Type + '\n');
              builder.push('' + '\n');
            }
            // Check if the Paragraph is an inserted revision.
            else if (paragraph.IsInsertRevision) {
              // Append information about the inserted revision to the builder.
              builder.push('The section' + document.GetIndex(section) + 'paragraph' + section.GetIndex(paragraph) + 'has been changed (inserted).' + '\n');
              builder.push('Author: ' + paragraph.InsertRevision.Author + '\n');
              builder.push('DateTime: ' + paragraph.InsertRevision.DateTime.ToString() + '\n');
              builder.push('Type: ' + paragraph.InsertRevision.Type + '\n');
              builder.push('' + '\n');
            }
            // Iterate over the child DocumentObjects in the Paragraph.
            else {
              for (let i = 0; i < paragraph.ChildObjects.Count; i++) {
                let obj = paragraph.ChildObjects.get(i);
                // Check if the child DocumentObject is a TextRange.
                if (obj.DocumentObjectType == wasmModule.DocumentObjectType.TextRange) {
                  let textRange = obj;
                  {
                    // Check if the TextRange is a deleted revision.
                    if (textRange.IsDeleteRevision) {
                      builder.push(
                        'The section' + document.GetIndex(section) + 'paragraph' + section.GetIndex(paragraph) + 'textrange' + paragraph.GetIndex(textRange) + 'has been changed (deleted).' + '\n'
                      );
                      builder.push('Author: ' + textRange.DeleteRevision.Author + '\n');
                      builder.push('DateTime: ' + textRange.DeleteRevision.DateTime.ToString() + '\n');
                      builder.push('Type: ' + textRange.DeleteRevision.Type + '\n');
                      builder.push('Change Text: ' + textRange.Text + '\n');
                      builder.push('' + '\n');
                    }
                    // Check if the TextRange is an inserted revision.
                    else if (textRange.IsInsertRevision) {
                      builder.push(
                        'The section' + document.GetIndex(section) + 'paragraph' + section.GetIndex(paragraph) + 'textrange' + paragraph.GetIndex(textRange) + 'has been changed (deleted).' + '\n'
                      );
                      builder.push('Author: ' + textRange.InsertRevision.Author + '\n');
                      builder.push('DateTime: ' + textRange.InsertRevision.DateTime.ToString() + '\n');
                      builder.push('Type: ' + textRange.InsertRevision.Type + '\n');
                      builder.push('Change Text: ' + textRange.Text + '\n');
                      builder.push('' + '\n');
                    }
                  }
                }
              }
            }
          }
        }
        // Define the output file name
        const outputFileName = 'GetParagraphRevisionsDetails.txt';
        // Save result file
        wasmModule.FS.writeFile(outputFileName, builder.join(''));

        // Read the saved file and convert to a Blob object
        const modifiedFileArray = wasmModule.FS.readFile(outputFileName);
        const modifiedFile = new Blob([modifiedFileArray], {
          type: 'text/plain',
        });

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
