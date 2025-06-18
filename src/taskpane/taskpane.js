Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {

    const paragraph = context.document.body.insertParagraph("gsjufsdhudsuifds发国后dhf", Word.InsertLocation.end);
    const wjj = context.document.body.insertParagraph("jk搞fsi规划风格jklhjkh发部署", Word.InsertLocation.end);
    const wkk = context.document.body.insertParagraph("sd看esduisddf", Word.InsertLocation.end);

    paragraph.font.color = "red";
    wjj.font.color = "pink";
    wkk.font.color = "blue";
    await context.sync();
  });
}
