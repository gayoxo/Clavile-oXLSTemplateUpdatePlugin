/**
 * 
 */
package fdi.ucm.server.updateparser.xlstemplate;

import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;

import fdi.ucm.server.modelComplete.collection.CompleteCollection;
import fdi.ucm.server.modelComplete.collection.document.CompleteDocuments;
import fdi.ucm.server.modelComplete.collection.document.CompleteTextElement;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteGrammar;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteStructure;
import fdi.ucm.server.modelComplete.collection.grammar.CompleteTextElementType;
import fdi.ucm.server.updateparser.xlstemplate.struture.Hoja;
import fdi.ucm.server.updateparser.xlstemplate.struture.HojaAntigua;
import fdi.ucm.server.updateparser.xlstemplate.struture.HojaNueva;

import java.io.FileInputStream;
import java.util.Iterator;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Clase que implementa la creacion de la base de datos per se
 * @author Joaquin Gayoso-Cabada
 *
 */
public class CollectionXLSTemplate implements InterfaceXLSTemplateparser {


	private static final String XLS_COLLECTION = "XLS COllection";
	private static final String COLECCION_A_APARTIR_DE_UN_XLS = "Coleccion a partir de un XLS";
	private CompleteCollection coleccionstatica;
	
	public CollectionXLSTemplate() {
		coleccionstatica=new CompleteCollection(XLS_COLLECTION, COLECCION_A_APARTIR_DE_UN_XLS+ new Timestamp(new Date().getTime()));
	}
	
	/* (non-Javadoc)
	 * @see fdi.ucm.server.importparser.sql.SQLparserModel#ProcessAttributes()
	 */
	@Override
	public void ProcessAttributes() {
		
	}

	
	 /**
	 
	  * Este metodo es usado para leer archivos Excel
	 
	  *
	 
	  * @param Nombre_Archivo
	 
	  *            - Nombre de archivo Excel.
	 
	  */
	 public void Leer_Archivo_Excel(String Nombre_Archivo) {
	 
	  /**
	 
	   * Crea una nueva instancia de Lista_Datos_Celda
	 
	   */
	 
	  ArrayList<Hoja> Hojas=new ArrayList<Hoja>();

	 
	  if (Nombre_Archivo.contains(".xlsx")) {
	 
		  Hojas=GENERAR_XLSX(Nombre_Archivo);
	 
	  } else if (Nombre_Archivo.contains(".xls")) {
	 
		  Hojas=GENERAR_XLS(Nombre_Archivo);

	 
	  }
	 
	  /**
	 
	   * Llama el metodo Imprimir_Consola para imprimir los datos de la celda
	 
	   * en la consola.
	 
	   */
	 
	  Imprimir_Consola(Hojas);
	 
	 }
	 
	 private ArrayList<Hoja> GENERAR_XLSX(String Nombre_Archivo) {
	 
		 ArrayList<Hoja> Salida=new ArrayList<Hoja>();
		 
	  try {
	 
	   /**
	 
	    * Crea una nueva instancia de la clase FileInputStream
	 
	    */
	 
	   FileInputStream fileInputStream = new FileInputStream(
	 
	     Nombre_Archivo);
	 
	   /**
	 
	    * Crea una nueva instancia de la clase XSSFWorkBook
	 
	    */
	 
	   XSSFWorkbook Libro_trabajo = new XSSFWorkbook(fileInputStream);
	 
	   
	   
	   int NStilos=Libro_trabajo.getNumberOfSheets();
		 
	   for (int i = 0; i < NStilos; i++) {
		   XSSFSheet Hoja_hssf = Libro_trabajo.getSheetAt(i);
		   HojaNueva Hojax=new HojaNueva(Hoja_hssf.getSheetName());
		   
		   Iterator<Row> Iterador_de_Fila = Hoja_hssf.rowIterator();
			 
		   List<List<XSSFCell>> Lista_Datos_Celda2 = new ArrayList<List<XSSFCell>>();
		   
		   while (Iterador_de_Fila.hasNext()) {
		 
			   XSSFRow Fila_hssf = (XSSFRow) Iterador_de_Fila.next();
		 
		    Iterator<Cell> iterador = Fila_hssf.cellIterator();
		 
		    List<XSSFCell> Lista_celda_temporal = new ArrayList<XSSFCell>();
		 
		    while (iterador.hasNext()) {
		 
		    	XSSFCell Celda_hssf = (XSSFCell) iterador.next();
		 
		     Lista_celda_temporal.add(Celda_hssf);
		 
		    }
		 
		    Lista_Datos_Celda2.add(Lista_celda_temporal);
		 
		   }
		   
		   Hojax.setListaHijos(Lista_Datos_Celda2);
		   Salida.add(Hojax);
	}
	   
	   

	  } catch (Exception e) {
	 
	   e.printStackTrace();
	 
	  }
	 
	  return Salida;
	 }
	 
	 private ArrayList<Hoja> GENERAR_XLS(String Nombre_Archivo) {
	 
		 ArrayList<Hoja> Salida=new ArrayList<Hoja>();
		 
	  try {
	 
	   /**
	 
	    * Crea una nueva instancia de la clase FileInputStream
	 
	    */
	 
	   FileInputStream fileInputStream = new FileInputStream(
	 
	     Nombre_Archivo);
	 
	   /**
	 
	    * Crea una nueva instancia de la clase POIFSFileSystem
	 
	    */
	 
	   POIFSFileSystem fsFileSystem = new POIFSFileSystem(fileInputStream);
	 
	   /**
	 
	    * Crea una nueva instancia de la clase HSSFWorkBook
	 
	    */
	 
	   HSSFWorkbook Libro_trabajo = new HSSFWorkbook(fsFileSystem);
	   
	   int NStilos=Libro_trabajo.getNumberOfSheets();
	 
	   for (int i = 0; i < NStilos; i++) {
		   HSSFSheet Hoja_hssf = Libro_trabajo.getSheetAt(i);
		   HojaAntigua Hojax=new HojaAntigua(Hoja_hssf.getSheetName());
		   
		   Iterator<Row> Iterador_de_Fila = Hoja_hssf.rowIterator();
			 
		   List<List<HSSFCell>> Lista_Datos_Celda2 = new ArrayList<List<HSSFCell>>();
		   
		   while (Iterador_de_Fila.hasNext()) {
		 
		    HSSFRow Fila_hssf = (HSSFRow) Iterador_de_Fila.next();
		 
		    Iterator<Cell> iterador = Fila_hssf.cellIterator();
		 
		    List<HSSFCell> Lista_celda_temporal = new ArrayList<HSSFCell>();
		 
		    while (iterador.hasNext()) {
		 
		     HSSFCell Celda_hssf = (HSSFCell) iterador.next();
		 
		     Lista_celda_temporal.add(Celda_hssf);
		 
		    }
		 
		    Lista_Datos_Celda2.add(Lista_celda_temporal);
		 
		   }
		   
		   Hojax.setListaHijos(Lista_Datos_Celda2);
		   Salida.add(Hojax);
	}
	   
	   
	  } catch (Exception e) {
	 
	   e.printStackTrace();
	 
	  }
	  
	  return Salida;
	 
	 }
	 
	 /**
	 
	  * Este método se utiliza para imprimir los datos de la celda a la consola.
	 
	  *
	 
	  * @param Datos_celdas
	 
	  *            - Listado de los datos que hay en la hoja de cálculo.
	 
	  */
	 
	 private void Imprimir_Consola(List<Hoja> HojasEntrada) {
	 
		 Long counterbase=-1l;
	
	for (Hoja hoja : HojasEntrada) {
		
		System.out.println("Nombre: " + hoja.getName());
		
		CompleteGrammar Grammar=new CompleteGrammar(hoja.getName(), hoja.getName(), coleccionstatica);
		coleccionstatica.getMetamodelGrammar().add(Grammar);
		HashMap<Integer, CompleteTextElementType> Hash=new HashMap<Integer, CompleteTextElementType>();
		HashMap<String, CompleteTextElementType> HashPath=new HashMap<String, CompleteTextElementType>();
		if (hoja instanceof HojaAntigua)
		{
			
			List<List<HSSFCell>> Datos_celdas = ((HojaAntigua) hoja).getListaHijos();
			
			 String Valor_de_celda;
			 
			  for (int FilaX = 0; FilaX < Datos_celdas.size(); FilaX++) {
			 
				CompleteDocuments Doc=new CompleteDocuments(coleccionstatica, Grammar, "", "");  
				if (FilaX!=0&&FilaX!=1)
					coleccionstatica.getEstructuras().add(Doc);
				  
			   List<HSSFCell> Lista_celda_temporal = Datos_celdas.get(FilaX);
			 
			   Integer Ambito=0;
			   for (int ColumnaX = 0; ColumnaX < Lista_celda_temporal.size(); ColumnaX++) {
			 
			 
			     HSSFCell hssfCell = Lista_celda_temporal.get(ColumnaX);
			 
			     Valor_de_celda = hssfCell.toString();
			 
			     if (ColumnaX==0)
			     {
			    	 Doc.setDescriptionText(Valor_de_celda);
			    	 try {
			    		 Doc.setClavilenoid(Long.parseLong(Valor_de_celda));
						} catch (Exception e) {
							Doc.setClavilenoid(counterbase);
							counterbase--;
						}
			    	
			     }
			     else if (ColumnaX==1)
			     {
			    	 Doc.setDescriptionText(Valor_de_celda);

			    	
			     }
			     else
			     {
			    if (FilaX==0)
			    	 {
			    	CompleteTextElementType C=generaStructura(Valor_de_celda,Grammar,HashPath);
			    	Hash.put(new Integer(ColumnaX), C);
			    	System.out.print("Columna:" + Valor_de_celda + "\t\t");
			    	
			    	 }
			    else if (FilaX==1)
		    	 {
			    	CompleteTextElementType C=Hash.get(new Integer(ColumnaX));
			    	try {
			    		C.setClavilenoid(Long.parseLong(Valor_de_celda));
					} catch (Exception e) {
						C.setClavilenoid(-1l);
					}
			    	
		    	Hash.put(new Integer(ColumnaX), C);
		    	System.out.print("Columna:" + Valor_de_celda + "\t\t");
		    	
		    	 }
			    else 
			    	{
			    	
			    	ArrayList<String> ValuesTotal=ProcessValue(Valor_de_celda);
			    	
			    	for (int i = 0; i < ValuesTotal.size(); i++) {
						CompleteTextElementType C=Hash.get(new Integer(ColumnaX));
				    	CompleteTextElement CT=new CompleteTextElement(C, ValuesTotal.get(i));
				    	ArrayList<Integer> ALIst = new ArrayList<Integer>();
				    			ALIst.add(Ambito);
				    	CT.setAmbitos(ALIst);
				    	Ambito=new Integer(Ambito+1);
				    	Doc.getDescription().add(CT);
				    	System.out.print("Valor:" + Valor_de_celda + "\t\t");
						
					}
			    	
			    	
			    	}
			      
			     }
			 
			   }
			 
			   System.out.println();
			 
			  }
		}
		else if (hoja instanceof HojaNueva)
		{
			
			List<List<XSSFCell>> Datos_celdas = ((HojaNueva) hoja).getListaHijos();
			
			 String Valor_de_celda;
			 
			  for (int FilaX = 0; FilaX < Datos_celdas.size(); FilaX++) {
			 
				  List<XSSFCell> Lista_celda_temporal = Datos_celdas.get(FilaX);
				  
				  CompleteDocuments Doc=new CompleteDocuments(coleccionstatica, Grammar, Integer.toString(FilaX), "");  
					if (FilaX!=0)
						coleccionstatica.getEstructuras().add(Doc);
			 
					Integer Ambito=0;
			   for (int ColumnaX = 0; ColumnaX < Lista_celda_temporal.size(); ColumnaX++) {
			 
			  
			 
			     XSSFCell hssfCell = (XSSFCell) Lista_celda_temporal.get(ColumnaX);
			 
			     Valor_de_celda = hssfCell.toString();
			 
			     if (ColumnaX==0)
			     {
			    	 Doc.setDescriptionText(Valor_de_celda);
			    	 try {
			    		 Doc.setClavilenoid(Long.parseLong(Valor_de_celda));
						} catch (Exception e) {
							Doc.setClavilenoid(counterbase);
							counterbase--;
						}
			    	
			     }
			     else if (ColumnaX==1)
			     {
			    	 Doc.setDescriptionText(Valor_de_celda);

			    	
			     }
			     else
			     
			     {
			    if (FilaX==0)
			    	 {
			    	CompleteTextElementType C=generaStructura(Valor_de_celda,Grammar,HashPath);
			    	Hash.put(new Integer(ColumnaX), C);
			    	System.out.print("Columna:" + Valor_de_celda + "\t\t");
			    	
			    	 }
			    else if (FilaX==1)
		    	 {
			    	CompleteTextElementType C=Hash.get(new Integer(ColumnaX));
			    	try {
			    		C.setClavilenoid(Long.parseLong(Valor_de_celda));
					} catch (Exception e) {
						C.setClavilenoid(-1l);
					}
			    	
		    	Hash.put(new Integer(ColumnaX), C);
		    	System.out.print("Columna:" + Valor_de_celda + "\t\t");
		    	
		    	 }
			    else 
			    	{
			    	ArrayList<String> ValuesTotal=ProcessValue(Valor_de_celda);
			    	
			    	for (int i = 0; i < ValuesTotal.size(); i++) {
						CompleteTextElementType C=Hash.get(new Integer(ColumnaX));
				    	CompleteTextElement CT=new CompleteTextElement(C, ValuesTotal.get(i));
				    	ArrayList<Integer> ALIst = new ArrayList<Integer>();
				    			ALIst.add(Ambito);
				    	CT.setAmbitos(ALIst);
				    	Ambito=new Integer(Ambito+1);
				    	Doc.getDescription().add(CT);
				    	System.out.print("Valor:" + Valor_de_celda + "\t\t");
						
					}
			    	
			    	}
			 
			   }
			 
			   System.out.println();
			 
			  }
			  }
		}
		
		
		
		
		
	}	 
		 
		 
	 
	 
	 }
	 
	 private ArrayList<String> ProcessValue(String valor_de_celda) {
		 Pattern regex = Pattern.compile("^Size:\\d+ (\\{\\w+\\})+$");
		 Matcher matcher = regex.matcher(valor_de_celda);
		 ArrayList<String> Salida=new ArrayList<String>();
		 if (matcher.matches())
		 {
			 String Temp=valor_de_celda.replace('{', ' ');
			 String[] list=Temp.split("}");
			 Long it=Long.parseLong(list[0].substring(5));
			 for (int i = 0; i < it; i++) {
				String E = list[i+1];
				E=E.trim();
				Salida.add(E);
			}
			 
		 }else
		 {
			 Salida.add(valor_de_celda);
		 }		 
		return Salida;
	}

	private CompleteTextElementType generaStructura(String valor_de_celda, CompleteGrammar grammar, HashMap<String, CompleteTextElementType> hashPath) {
		 
		
		 CompleteTextElementType preproducido = hashPath.get(valor_de_celda);
			if (preproducido!=null)
				return preproducido;
		 
		 
		String[] pathL=valor_de_celda.split("\\\\");
		
		CompleteStructure Padre=null;
		
		 if (pathL.length>1)
			 Padre=producePadre(pathL,hashPath,grammar);
		 
		 CompleteTextElementType Salida=null;
		if (Padre!=null)
		 {
			Salida=new CompleteTextElementType(pathL[pathL.length-1], Padre);
			Padre.getSons().add(Salida);
		 }
		else 
			{
			Salida=new CompleteTextElementType(valor_de_celda, grammar);
			grammar.getSons().add(Salida);
			}
		
		hashPath.put(valor_de_celda, Salida);
		return Salida;
	}

	private CompleteStructure producePadre(String[] pathL,
			HashMap<String, CompleteTextElementType> hashPath,CompleteGrammar CG) {
		
		String Acumulado = "";
		CompleteTextElementType Padre = null;
		for (int i = 0; i < pathL.length-1; i++) {
			if (i!=0)
				Acumulado=Acumulado+"\\"+pathL[i];
			else
				Acumulado=Acumulado+pathL[i];
			CompleteTextElementType yo = hashPath.get(Acumulado);
			if (yo==null)
				{
				
				if (Padre!=null)
					{
					CompleteTextElementType Salida = new CompleteTextElementType(pathL[i], Padre);
					Padre.getSons().add(Salida);
					hashPath.put(Acumulado, Salida);
					}
				else
					{
					CompleteTextElementType Salida = new CompleteTextElementType(pathL[i], CG);
					CG.getSons().add(Salida);
					hashPath.put(Acumulado, Salida);
					}
				
				}
			
			Padre=yo;
		}
		return Padre;
	}

	public static void main(String[] args) {
		
		main2();

	  String fileName = "ejemplo2.xls";
	 
	  System.out.println(fileName);
	 
	 CollectionXLSTemplate C = new CollectionXLSTemplate();
	 C.Leer_Archivo_Excel(fileName);
	 C.toString();
	 }

	@Override
	public void ProcessInstances() {
		
		
	}


	public CompleteCollection getColeccion() {
		return coleccionstatica;
	}
	
	public static void main2()
	{
		//^Size:d+(\\{(w+)\\})+$
		Pattern regex = Pattern.compile("^Size:\\d+ (\\{\\w+\\})+$");
		 Matcher matcher = regex.matcher("Size:55 {aaaa}{dd}{ccc}");
		 System.out.println(matcher.matches());

	}
}
