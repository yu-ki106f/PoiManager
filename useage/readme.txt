テンプレートの利用、一括入力、一括取込
1.「PoiManager.xla」をExcel Addinに追加
2.PoiManager -> Excel -> Cells setting
   ・ セルの入出力設定
3.PoiManager -> Excel -> Make XML
   ・入出力設定XML作成（クリップボード経由）
4.Java Project resourceフォルダにXMLファイルを設置
5.テンプレートExcelと生成XMLを読込み、データを流し込んで別名保存
  PoiManager.getInstance(new File("テンプレート.xlsx")).getReadWriter()
            .loadFromXml("生成した.xml")
            .write(InputDataDto)
            .saveBook(new File("保存.xlsx"));
6.入力済みExcelからデータを読込み
  InputDataDto dto = PoiManager.getInstance(new File("入力済みExcel.xlsx")).getReadWriter()
                     .loadFromXml("生成した.xml")
                     .read(InputDataDto.class);


流れるようなインターフェースを意識した実装
クラス名と役割が適当なきがします・・・。

Ex)
PoiManager.getInstance(new File("test1.xls"))
.sheet("main")			//The "main" sheet is chosen. It is created when it does not exist. 
	.range("B2:G2")
		.style()
			.border().all().type().BORDER_MEDIUM()
			.border().all().color().BLACK()
			.fill().foregroundColor().SKY_BLUE()
			.fill().pattern().SOLID_FOREGROUND()
			.font().BOLD()
			.font().points(12)
			.font().color().BLACK()
			.font().name("ＭＳ ゴシック")
		.update()							//style update
		.setColumnWidth(15)					//column size 15
	.range("B3:G20")
		.comment("Test").write()			//comment append
		.style()
			.border().all().type().BORDER_MEDIUM()
			.border().all().color().BLUE_GREY()
			.update()
	.cell("B2").setValue(new Date()).
			style().setDataFormat("yyyy/mm/dd")
				.update()
			.comment("コメント").write()
	.cell("B1").comment("コメントです").write()
.saveBook();								//save

