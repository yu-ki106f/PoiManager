package operation;

import java.io.File;
import java.util.Date;

import org.poco.framework.poi.managers.IReadWriter;
import org.poco.framework.poi.managers.IPoiManager.IPoiBook;
import org.poco.framework.poi.managers.impl.PoiManager;
import org.poco.framework.poi.utils.WorkbookUtil;
import operation.dto.SampleParentDto;
import operation.utils.PosTestUtil;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 
 */
public class PoiTest {

	/**
	 * 流れるようなインターフェースによるPOI操作実装テスト
	 * @param args
	 */
	public static void main(String[] args) {
		//テンプレートファイルまたは、存在しないファイルでも可能（自動的に新規ブックを作成する）
		//存在しないファイルの場合は、saveBookされるまでは、ファイルは生成されない。
		//Fileを内部管理のハッシュキーに使用しているだけなので、なくても必要。
		
		//メソッド操作による内容更新（値、スタイル） → 途中保存 → イメージ追加等 → 保存
		//Sheet,Range,Merge,Cellから、内容スタイル更新
		test1CreateExcel(new File("output/test1Create.xls"));	//xls
		test1CreateExcel(new File("output/test1Create.xlsx"));	//xlsx
		
		//メソッド操作による内容更新 → コメント追加 → 保存
		test2CreateXlsOrXlsx(new File("output/test2Create.xls"));	//xls
		test2CreateXlsOrXlsx(new File("output/test2Create.xlsx"));	//xlsx
		
		//書き込み、読み込み設定ファイルであるXMLは、FLFramework依存しない為、key,fileName,path属性は利用せず、
		//PoiManagerインスタンスをそのまま利用する形で動作する
		
		//Excelアドインで生成されるxmlデータを利用して、オブジェクトデータからExcelに内容書き込み
		test3WriteDataExcel(new File("template/Sample.xls"),new File("output/test3Write.xls"),500,250);	//xls
		test3WriteDataExcel(new File("template/Sample.xlsx"),new File("output/test3Write.xlsx"),500,250);	//xlsx
		
		//Excelアドインで生成されるxmlデータを利用して、Excelファイルからオブジェクトデータを読込む
		//(bookの属性name.key,pathは読込まないため、xls,xlsxどちらのフォーマットも同一ファイルで動作可能）
		test4ReadDataXls(new File("output/test3Write.xls"),250);		//xls	test3WriteDataExcelで作成したExcelファイルを利用
		test4ReadDataXls(new File("output/test3Write.xlsx"),500);	//xlsx	test3WriteDataExcelで作成したExcelファイルを利用

		//5万件
		test3WriteDataExcel(new File("template/Sample.xls"),new File("output/test3Write50000.xls"),50000,30000);	//xls
		test3WriteDataExcel(new File("template/Sample.xlsx"),new File("output/test3Write50000.xlsx"),50000,30000);	//xlsx

	}

	/**
	 * 指定されたファイル拡張子（新規作成時）によって、生成するエクセルフォーマットを自動判定し、Excelファイルを生成する
	 * @param f
	 */
	private static void test1CreateExcel(File f) {
		
		System.out.println("START " + getCallMethodName() + "_" + f.getName());
		try {
			//Fileクラス単位で内部で管理している為、分割呼出ししても問題なく継続される
			PoiManager.getInstance(f).sheet("シート").cell("A1").setValue("文字を書く");
			PoiManager.getInstance(f).sheet("シート").cell("A2").setValue("文字を書く");
			PoiManager.getInstance(f).sheet("シート").cell(1, 2).setValue(new Date());
			//型によって格納するタイプを自動判定
			PoiManager.getInstance(f).sheet("シート_N").cell(1, 2).setValue("123");	//文字列で格納
			PoiManager.getInstance(f).sheet("シート_N").cell(1, 3).setValue(123);		//数値で格納
			//normal　数値で格納
			PoiManager.getInstance(f).sheet("シート_N").cell(1, 2).getOrgCell().setCellValue(123);
			PoiManager.getInstance(f).sheet("シート_N").cell(1, 2).getOrgCell().setCellType(Cell.CELL_TYPE_NUMERIC);
			
			PoiManager.getInstance(f).sheet("シート")
				.cell("B1")
					.style()
						.align().ALIGN_CENTER()
						.border().all().type().BORDER_MEDIUM()
						.fill().foregroundColor().PINK()
						.fill().pattern().SOLID_FOREGROUND()
						.font().BOLD()
						.font().name("HGS創英角ｺﾞｼｯｸUB")
					.update()
					.setValue("文字各少し眺め").autoSizeColumn()
				//セルの結合
				.merge("20:21").setValue("20-21行")
					.style()
						.valign().VERTICAL_TOP()
						.font().points(15)
					.update().setRowWidth(10).setRownHidden(true)
					.setRownHidden(false)
				//行の追加
				.insertRows("12:14").setValue("Insert12-14").setRowWidth(12.75);
			
			
			PoiManager.getInstance(f).sheet("シート").cell("B3")
			.style().font().name("HGS創英角ｺﾞｼｯｸUB").font().BOLD()
			.update()
			.saveBook();	//testCreateXls1.xlsで保存
			
			//シートをシート２としてクローン作成し、ソノシートの値やスタイルを更新しつつ、保存する
			PoiManager.getInstance(f).cloneSheet("シート", "シート2")
			.range("C2:Z100")
				.style()
					.border().all().type().BORDER_DASHED()
					.align().ALIGN_CENTER()
					.fill().foregroundColor().GOLD()
					.fill().pattern().SOLID_FOREGROUND()
					.font().color().RED()
					.font().name("HGP創英角ﾎﾟｯﾌﾟ体")
					.font().points((short)12)
					.update()
					.setValue("文字")
			.merge("A3:A8")
				.style()
					.fill().foregroundColor().AQUA()
					.fill().pattern().SOLID_FOREGROUND()
					.font().color().GREEN()
					.font().points((short)15)
					.font().BOLD()
					.valign().VERTICAL_CENTER()
					.update()
					.setValue("結合1")
			.merge("B2:C3")
				.style()
					.valign().VERTICAL_TOP()
					.update()
				.setValue("結合")
			.range("B5:B6")
				//画像設定
				.setImage(new File("template/image.png") )
					.padding(250, 10, 250, 10)
				.drow()
			.range("C1:Z1")
				.style()
					.align().ALIGN_RIGHT()
					.update()
				.setValue("右寄せ")
				.selected()
			//test1Create-1.xls or xlsxで保存
			.saveBook(new File("output/test1Create-1." + getSuffix(f.getName())));
		
		}
		catch (Exception e){
			e.printStackTrace();
		}
		finally {
			PoiManager.destroy();
		}
		System.out.println("END " + getCallMethodName() + "_" + f.getName());
	}
	/**
	 * 指定されたファイルを新規作成し、値やスタイルをメソッド呼び出しにより書き込み、保存を行う
	 */
	private static void test2CreateXlsOrXlsx(File f) {
		System.out.println("START " + getCallMethodName() + "_" + f.getName());
		try {
			//内容を書き込んで保存
			test2Create(PoiManager.getInstance(f)).saveBook();
			//Date型での取得も可能
			Date date =PoiManager.getInstance(f).sheet("メイン").cell("B2").value().getDate();
			System.out.println(date);
		}
		catch(Exception e){
			e.printStackTrace();
		}
		finally {
			PoiManager.destroy();
		}
		System.out.println("END " + getCallMethodName() + "_" + f.getName());
	}
	/**
	 * bookに内容を書き込む
	 * @param book
	 * @return
	 */
	private static IPoiBook test2Create(IPoiBook book) {
		return book.sheet("メイン")
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
			.update()
			.setColumnWidth(15)
		.range("B3:G20")
			.comment("テスト").write()
			.style()
				.border().all().type().BORDER_MEDIUM()
				.border().all().color().BLUE_GREY()
				.update()
		.cell("B2").setValue(new Date()).
				style().setDataFormat("yyyy/mm/dd")
					.update()
				.comment("コメント").write()
				.cell("B1").comment("コメントです").write();
	}
	
	/**
	 * 指定されたテンプレートExcelに指定サイズ分のデータを流し込み、別名保存を行う<BR/>
	 * 注意：引数なしのSaveBook()呼び出しでテンプレートが書き換わってしまうため、別名保存を行うこと
	 * 
	 * ※1 XLSXは存在しない行を列にスタイルを持たせることで、テンプレート側のスタイルが適応されるような
	 * 設定が有効にはならず、新規Cellとなってしまう為、先頭行のスタイルをコピー機能で対応。
	 * ただし、大量データ出力時は、大幅にパフォーマンスが低下する為、非推奨
	 */
	private static void test3WriteDataExcel(File f, File out, Integer size, Integer limit) {
		System.out.println("START " + getCallMethodName() + "_" + f.getName() + "_50000");
		
		try {
			IPoiBook book =	PoiManager.getInstance(f);
			boolean isXlsx = WorkbookUtil.isXlsx(book.getOrgWorkBook());
			//10万件のデータ
			book.getReadWriter()
			.loadFromXml("Sample.xml")
			.setMaxRows(limit)			//指定件数でシート分割(XLSXは無視)
			.setStyleCopy(isXlsx)		//※1
			.write(PosTestUtil.createData(size))
			.saveBook(out);
		}
		catch(Exception e){
			e.printStackTrace();
		}
		finally {
			PoiManager.destroy();
		}
		System.out.println("END " + getCallMethodName() + "_" + f.getName() + "_50000");
	}
	
	/**
	 * testWriteDataXlsで作成したoutput.xlsファイルを読込み、<BR/>
	 * エクセルファイルからデータを取得し、生成データと一致しているかの確認を行う
	 */
	private static void test4ReadDataXls(File f,Integer size) {
		System.out.println("START " + getCallMethodName() + "_" + f.getName() + "_50000");

		try {
			//シートは設定ファイルでは1つのため、分割した5万件までしか取り込めない
			IPoiBook book =	PoiManager.getInstance(f);
			IReadWriter w = book.getReadWriter();
			SampleParentDto result = w.loadFromXml("Sample.xml")
										.read(SampleParentDto.class);
			
			System.out.println("read data count:" + result.list.size());
			//読込んだリストと生成したリスト値の比較を行う
			Boolean isEqual = PosTestUtil.equals(
								PosTestUtil.createData(size).list,
								result.list);
			System.out.println("contents equals :" + isEqual);
		}
		catch(Exception e){
			e.printStackTrace();
		}
		finally {
			PoiManager.destroy();
		}
		System.out.println("END " + getCallMethodName() + "_" + f.getName() + "_50000");
	}
	
	/**
	 * 呼び出しメソッド名を返す
	 * @return
	 */
	private static String getCallMethodName() {
		Exception e = new Exception();
		return e.getStackTrace()[1].getMethodName();
	}
	/**
	 * ファイル名から拡張子を返します。
	 * 
	 * @param fileName
	 *            ファイル名
	 * @return ファイルの拡張子
	 */
	private static String getSuffix(String fileName) {
		if (fileName == null)
			return null;
		int point = fileName.lastIndexOf(".");
		if (point != -1) {
			return fileName.substring(point + 1);
		}
		return "";
	}
}
