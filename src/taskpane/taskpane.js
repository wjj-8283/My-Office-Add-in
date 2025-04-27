Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {

    const paragraph = context.document.body.insertParagraph("gsjuf机设计和发国后dhf", Word.InsertLocation.end);
    const wjj = context.document.body.insertParagraph("jk搞fsi规划风格化风格化风格化jklhjkh发部署", Word.InsertLocation.end);
    const wff = context.document.body.insertParagraph("cb就开会开会就开会风格恢复书gfdhg", Word.InsertLocation.end);
    const wkk = context.document.body.insertParagraph("sd看gujfg是对uihguisddf", Word.InsertLocation.end);

    paragraph.font.color = "purple";
    wjj.font.color = "pink";
    wff.font.color = "red";
    wkk.font.color = "blue";
    await context.sync();
  });
}
