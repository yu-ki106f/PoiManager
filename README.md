PoiManager
==========

###It is an Excel generation support framework using Apache Poi.

### xls, xlsx support.

####Checking of version poi 3-10-x

####Sample Code
	PoiManager.getInstance(new File("test1.xlsx"))    //The "test1.xlsx" file is read. It is created when it does not exist.
	.sheet("main")                                    //The "main" sheet is chosen. It is created when it does not exist.
	    .range("B2:G2")
	        .style()
	            .border().all().type().BORDER_MEDIUM()
	            .border().all().color().BLACK()
	            .fill().foregroundColor().SKY_BLUE()
	            .fill().pattern().SOLID_FOREGROUND()
	            .font().BOLD().font().points(12).font().color().BLACK()
	            .font().name("ＭＳ ゴシック")
	        .update()                            //style update
	        .setColumnWidth(15)                  //column size 15
	    .range("B3:G20")
	        .comment("Test").write()             //comment append
	        .style()
	            .border().all().type().BORDER_MEDIUM()
	            .border().all().color().BLUE_GREY()
	        .update()
	    .cell("B2").setValue(new Date())
	        .style().setDataFormat("yyyy/mm/dd").update()
	        .comment("コメント").write()
	    .cell("B1").comment("コメントです").write()
	.saveBook();                                //save

	PoiManager.getInstance(new File("test1.xls"))    //The "test1.xls" file is read. It is created when it does not exist.
	.sheet("main")                                    //The "main" sheet is chosen. It is created when it does not exist.
	    .range("B2:G2")
	        .style()
	            .border().all().type().BORDER_MEDIUM()
	            .border().all().color().BLACK()
	            .fill().foregroundColor().SKY_BLUE()
	            .fill().pattern().SOLID_FOREGROUND()
	            .font().BOLD().font().points(12).font().color().BLACK()
	            .font().name("ＭＳ ゴシック")
	        .update()                            //style update
	        .setColumnWidth(15)                  //column size 15
	    .range("B3:G20")
	        .comment("Test").write()             //comment append
	        .style()
	            .border().all().type().BORDER_MEDIUM()
	            .border().all().color().BLUE_GREY()
	        .update()
	    .cell("B2").setValue(new Date())
	        .style().setDataFormat("yyyy/mm/dd").update()
	        .comment("コメント").write()
	    .cell("B1").comment("コメントです").write()
	.saveBook();                                //save

####Sample Code
	//write data
	PoiManager.getInstance(new File("template.xlsx")).getReadWriter()
            .loadFromXml("setting.xml")		//Please be created using the PoiManager.xla
            .write(inputDataDto)
            .saveBook(new File("output.xlsx"));

	PoiManager.getInstance(new File("template.xls")).getReadWriter()
            .loadFromXml("setting.xml")		//Please be created using the PoiManager.xla
            .setMaxRows(65000)
            .write(inputDataDto)
            .saveBook(new File("output.xls"));

####Sample Code
	//read data
	LoadData data = PoiManager.getInstance(new File("input.xlsx")).getReadWriter()
                  .loadFromXml("setting.xml")		//Please be created using the PoiManager.xla
                  .read(LoadData.class);

	LoadData data = PoiManager.getInstance(new File("input.xls")).getReadWriter()
                  .loadFromXml("setting.xml")		//Please be created using the PoiManager.xla
                  .read(LoadData.class);

####Sample Code
	//Sample Map list data.
	List<Object> mapList = createMapList();
	//Sample Dto list data.
	List<Object> dtoList = createDtoList();

	IPoiSheet sheet = PoiManager.getInstance(new File("writeObject.xlsx")).sheet("TestSheet");

	//Write to the vertical and horizontal than the starting position.

	//Writes the contents (column order specified)
	sheet.cell(0,7).writeObject(mapList).order(new String[]{"profit","cost","amount","name","product"}).write();
	//Writes the contents (column order specified)
	sheet.cell(6,7).writeObject(dtoList).order(new String[]{"name","amount","cost","profit"}).write();

	//No column order specified. append header.
	sheet.cell(12,6).writeObject(mapList).setHeader(createMap(true)).write()
		.cell(18,6).writeObject(dtoList).setHeader(createDto(true)).write()
		//save
		.saveBook();

## License
MIT
