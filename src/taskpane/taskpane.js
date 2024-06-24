Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {

    const paragraph = context.document.body.insertParagraph("bhd获得合法合规和雕塑覅呼呼大睡复古很好的李的飞机设计和发动机客户建国后dhf", Word.InsertLocation.end);
    const wjj = context.document.body.insertParagraph("jk搞撒低级刚才发图茂标的符合公司的开发部署", Word.InsertLocation.end);
    const wff = context.document.body.insertParagraph("cb书宽说的gfjvjvhbjv粉红色呃u匮乏昏聩的覅u是乎都是v福冈大阪dghdhg", Word.InsertLocation.end);
    const wkk = context.document.body.insertParagraph("sd看guts这个副sb业阿飞故园附庸国df", Word.InsertLocation.end);

    paragraph.font.color = "green";
    wjj.font.color = "orange";
    wff.font.color = "yellow";
    wkk.font.color = "blue";
    await context.sync();
  });
}
