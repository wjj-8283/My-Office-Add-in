Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {

    const paragraph = context.document.body.insertParagraph("bhdghrfufuyjfyju的bvdgfhg的飞机设计和发动机客户建国后dhf", Word.InsertLocation.end);
    const wjj = context.document.body.insertParagraph("jk搞fsdjhfjfysdj开jklhjkh发部署", Word.InsertLocation.end);
    const wff = context.document.body.insertParagraph("cb书宽说的gftytytjkdfhguioyhuiyiokjdf昏聩的覅u是乎都是v福冈大阪dghdhg", Word.InsertLocation.end);
    const wkk = context.document.body.insertParagraph("sd看gutggujfuftyuiyhiugs这个7807kui庸国df", Word.InsertLocation.end);

    paragraph.font.color = "orange";
    wjj.font.color = "pink";
    wff.font.color = "red";
    wkk.font.color = "blue";
    await context.sync();
  });
}
