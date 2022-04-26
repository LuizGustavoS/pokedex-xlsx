import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.JSONObject;

import javax.ws.rs.Consumes;
import javax.ws.rs.GET;
import javax.ws.rs.Path;
import javax.ws.rs.Produces;
import javax.ws.rs.core.MediaType;
import javax.ws.rs.core.Response;
import java.io.File;
import java.io.FileInputStream;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

@Path("/pokedex")
@Consumes(MediaType.APPLICATION_JSON)
@Produces(MediaType.APPLICATION_JSON)
public class LoaderController {

    @GET
    public Response test() throws Exception {

        final URL resource = getClass().getClassLoader().getResource("pokedex.xlsx");
        if (resource == null){
            throw new Exception("Arquivo xlsx não encontrado!");
        }

        //TODO melhorar essa importacao do arquivo
        final File file = new File(resource.toURI());
        final FileInputStream fis = new FileInputStream(file);

        XSSFWorkbook wb = new XSSFWorkbook(fis);
        XSSFSheet sheet = wb.getSheetAt(0);

        List<String> collumnsNames = new ArrayList<>();
        List<JSONObject> result = new ArrayList<>();

        boolean firstRow = true;
        for (Row row : sheet) {
            if (firstRow){
                collumnsNames = loadCollumnsNames(row);
                firstRow = false;
                continue;
            }

            result.add(loadRow(collumnsNames, row));
        }

        fis.close();
        return Response.ok(result).build();
    }

    //--

    private JSONObject loadRow(List<String> collumnsNames, Row row){

        JSONObject data = new JSONObject();

        final Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            final Cell cell = cellIterator.next();
            if (cell.getCellType().equals(CellType.STRING)) {
                data.put(collumnsNames.get(cell.getAddress().getColumn()), cell.getStringCellValue());
            } else if (cell.getCellType().equals(CellType.NUMERIC)) {
                data.put(collumnsNames.get(cell.getAddress().getColumn()), cell.getNumericCellValue());
            }
        }

        return data;
    }

    private List<String> loadCollumnsNames(Row row){

        List<String> listNames = new ArrayList<>();

        final Iterator<Cell> cellIterator = row.cellIterator();
        while (cellIterator.hasNext()) {
            final Cell cell = cellIterator.next();
            listNames.add(cell.getStringCellValue());
        }

        return listNames;
    }
}