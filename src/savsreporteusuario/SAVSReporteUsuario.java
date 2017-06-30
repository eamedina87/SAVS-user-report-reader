/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
package savsreporteusuario;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Map;
import java.util.TreeMap;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.CellType;
import org.xml.sax.SAXException;


/**
 *
 * @author erick.medina
 */
public class SAVSReporteUsuario {
    public static final String fileName = "D:\\Documentos\\Inspeccion AVS\\2017-05 ClaroTv\\Usuarios\\Base Diciembre.xls";
    public static final String JURIDICO_IDENTIFIER = "RUC";
    
    
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws InvalidFormatException, OpenXML4JException, SAXException {
         HSSFWorkbook wb = null;
        try {
           
            // TODO code application logic here
            wb = readFile(fileName);
            System.out.println("Data dump:\n");
          
            HSSFSheet sheet = wb.getSheetAt(0);
            int rows = sheet.getPhysicalNumberOfRows();
            ArrayList<Usuario> usuarios = new ArrayList<>();
            HashMap<String, String> ubicacion = new HashMap<>();
       
            for (int r = 1; r < rows; r++) {
                    HSSFRow row = sheet.getRow(r);
                    if (row == null) {
                            continue;
                    }
                                                         
                    Usuario usuario = new Usuario();       
                   
                    /*Formato del documento en Excel
                     * REGION - CANAL - PRODUCTO - SUBPRODUCTO - NOMBRE DE PLAN BASE - DECODIFICADORES - DATO_IDENTIFICACION
                     * TIPO IDENTIFICACION - DATO NOMBRE COMPLETO - CONCATENADO ARCOTEL - PROV ARCO - CIU ARCO - PARR ARCO
                     * */
                     
                    usuario.setREGION(row.getCell(0)!=null?row.getCell(0).getStringCellValue():"");
                    usuario.setCANAL(row.getCell(1)!=null?row.getCell(1).getStringCellValue():"");
                    usuario.setPRODUCTO(row.getCell(2)!=null?row.getCell(2).getStringCellValue():"");
                    usuario.setSUBPRODUCTO(row.getCell(3)!=null?row.getCell(3).getStringCellValue():"");
                    usuario.setNOMBRE_DE_PLAN_BASE(row.getCell(4)!=null?row.getCell(4).getStringCellValue():"");
                    usuario.setDECODIFICADORES(row.getCell(5)!=null?row.getCell(5).getNumericCellValue():0);
                    if (row.getCell(6)!=null && row.getCell(6).getCellTypeEnum()==CellType.STRING){
                        usuario.setDATO_IDENTIFICACION(row.getCell(6)!=null?row.getCell(6).getStringCellValue():"");
                    }
                    usuario.setTIPO_IDENTIFICACION(row.getCell(7)!=null?row.getCell(7).getStringCellValue():"");
                    usuario.setDATO_NOMBRE_COMPLETO(row.getCell(8)!=null?row.getCell(8).getStringCellValue():"");
                    usuario.setCONCATENADO_ARCOTEL(row.getCell(9)!=null?row.getCell(9).getStringCellValue():"");
                    String provincia = row.getCell(10)!=null?row.getCell(10).getStringCellValue():"";
                    usuario.setProv_Arco(provincia);
                    String ciudad = row.getCell(11)!=null?row.getCell(11).getStringCellValue():"";
                    usuario.setCiu_Arco(ciudad);
                    usuario.setParr_Arco(row.getCell(12)!=null?row.getCell(12).getStringCellValue():"");
                    
                    usuarios.add(usuario);
                          
                    ubicacion.put(ciudad, provincia);
                    
            }
            
                        
            TreeMap<String, String> mSortedUbicaciones = new TreeMap<>(ubicacion);
                       
            Iterator it = mSortedUbicaciones.entrySet().iterator();
            int mContTotal = 0;
            while (it.hasNext()) {
                Map.Entry pair = (Map.Entry)it.next();
                String mCiudad = (String) pair.getKey();
                String mProvincia = (String) pair.getValue();
                int mContJuridica = 0;
                int mContNatural = 0;
                
                for (Usuario usuario:usuarios){
                    if (!usuario.getCiu_Arco().equals(mCiudad)) continue;
                    mContTotal++;
                    if (usuario.getTIPO_IDENTIFICACION().equals(JURIDICO_IDENTIFIER)){
                        mContJuridica++;
                    } else {
                        mContNatural++;
                    }                 
                }
                
                System.out.println(mProvincia+";"+mCiudad+";"+mContJuridica+";"+mContNatural);
                
                it.remove(); // avoids a ConcurrentModificationException
            }
            
            System.out.println("Total:"+mContTotal);
           
        } catch (IOException ex) {
            Logger.getLogger(SAVSReporteUsuario.class.getName()).log(Level.SEVERE, null, ex);
        }  finally {
            try {
                wb.close();
            } catch (IOException ex) {
                Logger.getLogger(SAVSReporteUsuario.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }
    
    private static HSSFWorkbook readFile(String filename) throws IOException {
	    FileInputStream fis = new FileInputStream(filename);
	    try {
	        return new HSSFWorkbook(fis);		
	    } finally {
	        fis.close();
	    }
	}
}
