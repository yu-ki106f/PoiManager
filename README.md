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
	            .font().name("�l�r �S�V�b�N")
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
	        .comment("�R�����g").write()
	    .cell("B1").comment("�R�����g�ł�").write()
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
	            .font().name("�l�r �S�V�b�N")
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
	        .comment("�R�����g").write()
	    .cell("B1").comment("�R�����g�ł�").write()
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
