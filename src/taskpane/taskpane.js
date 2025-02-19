Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {

    const paragraph = context.document.body.insertParagraph("gsjuf机设计和发动机客户建国后dhf", Word.InsertLocation.end);
    const wjj = context.document.body.insertParagraph("jk搞fsjkdehuiughierdjhfgsgdddfhjkusxhfjjussdhhdgdhjdsjklhjkh发部署", Word.InsertLocation.end);
    const wff = context.document.body.insertParagraph("cb书gftyfygukfdufdtu运动发育东方影都都是浮云大阪dghdhg", Word.InsertLocation.end);
    const wkk = context.document.body.insertParagraph("sd看gutftshsjfghjksdhfdsjkhfyuitgf7kui庸国jf是对u士大夫以上的覅是对uihguisddf", Word.InsertLocation.end);

    paragraph.font.color = "purple";
    wjj.font.color = "pink";
    wff.font.color = "red";
    wkk.font.color = "blue";
    await context.sync();
  });
}
