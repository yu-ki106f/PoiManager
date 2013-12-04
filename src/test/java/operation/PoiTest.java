/**
 * 
 */
package operation;

import java.io.File;
import java.util.Date;

import org.poco.framework.poi.managers.impl.PoiManager;

import org.apache.poi.ss.usermodel.Cell;

/**
 * 
 * @author yu-ki106f
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
		try {
			testExcel1();
			testExcel2();
		}
		catch (Exception e){
			e.printStackTrace();
		}
		finally {

		}

	}
	
	public static void testExcel2() {
		File f = new File("C:\\test2.xls");

		try {
			PoiManager.getInstance(f)
			.sheet("メイン")
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
				.cell("B1").comment("コメントです").write()
			.saveBook();
			
			System.out.print(PoiManager.getInstance(f).sheet("メイン").cell("B2").value().getDate());
		}
		catch(Exception e){
			e.printStackTrace();
		}
		finally {
			PoiManager.destroy();
		}
	}
	
	public static void testExcel1() {
		
		File f = new File("C:\\test1.xls");
		
		try {
			
			//Fileクラス単位で内部で管理している為、分割呼出ししても問題なく継続される
			PoiManager.getInstance(f).sheet("シート").cell("A1").setValue("文字を書く");
			PoiManager.getInstance(f).sheet("シート").cell("A2").setValue("文字を書く");
			PoiManager.getInstance(f).sheet("シート").cell(1, 2).setValue(new Date());
			
			PoiManager.getInstance(f).sheet("シート_N").cell(1, 2).setValue("123");
			PoiManager.getInstance(f).sheet("シート_N").cell(1, 3).setValue(123);
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
			.saveBook();	//test1.xlsで保存
			
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
				.setImage(new File("c:\\logout.png") )
					.padding(250, 10, 250, 10)
				.drow()
			.range("C1:Z1")
				.style()
					.align().ALIGN_RIGHT()
					.update()
				.setValue("右寄せ")
				.selected()
			.saveBook(new File("c:\\test1-1.xls"));
			//test1-1.xlsで保存
		
		}
		catch (Exception e){
			e.printStackTrace();
		}
		finally {
			PoiManager.destroy();
		}
	}

}
