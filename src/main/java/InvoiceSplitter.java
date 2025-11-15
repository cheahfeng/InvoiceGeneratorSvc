import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfReader;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.kernel.pdf.canvas.parser.PdfTextExtractor;
import com.itextpdf.kernel.pdf.canvas.parser.listener.LocationTextExtractionStrategy;
import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class InvoiceSplitter {
    private static final Logger logger = LogManager.getLogger(InvoiceSplitter.class);

    public static void main(String[] args) throws IOException {
        InvoiceSplitter app = new InvoiceSplitter();
        app.run();
    }

    // service type sort order (used for 2nd file)
    private final Map<String, Integer> serviceTypeOrder = new LinkedHashMap<>();

    // totals from 2nd file: companyKey -> (serviceType -> amount)
    private final Map<String, Map<String, BigDecimal>> serviceTotalsByCompany = new LinkedHashMap<>();

    // CHSS totals from 1st file: companyKey -> amount
    private final Map<String, BigDecimal> chssTotalByCompany = new LinkedHashMap<>();

    // path to your Excel template
    private final String templatePath = "template/Template.xlsx";

    public InvoiceSplitter() {
        serviceTypeOrder.put("TAX", 1);
        serviceTypeOrder.put("ACCOUNT", 2);
        serviceTypeOrder.put("BPO", 3);
        serviceTypeOrder.put("SECRETARY", 4);
        serviceTypeOrder.put("OTHERS", 5);
    }

    public void run() throws IOException {
        System.out.println("Generation Started...");

        List<SourceConfig> sources = new ArrayList<>();

        sources.add(new SourceConfig(
                "input/INVOICE - CHSS.pdf",
                new InvoiceExtractor(),
                1,      // namePriority
                1,      // inputOrder
                false   // sortByService
        ));

        sources.add(new SourceConfig(
                "input/SOA - SHAREBIZ.pdf",
                new StatementOfAccountExtractor(),
                2,      // namePriority
                2,      // inputOrder
                false   // sortByService
        ));

        sources.add(new SourceConfig(
                "input/INVOICE - SHAREBIZ.pdf",
                new InvoiceByTypeExtractor(),
                3,      // namePriority
                3,      // inputOrder
                true    // sortByService
        ));

        String outputDir = "output/companies/";
        processInvoices(sources, outputDir);
        System.out.println("Generation Completed...");
    }

    private void processInvoices(List<SourceConfig> sources, String outputDir) throws IOException {
        File outDirFile = new File(outputDir);
        if (!outDirFile.exists()) {
            outDirFile.mkdirs();
        }

        // open all PDFs
        Map<String, PdfDocument> openDocs = new LinkedHashMap<>();
        for (SourceConfig src : sources) {
            openDocs.put(src.path, new PdfDocument(new PdfReader(src.path)));
        }

        // companyKey -> list of pages
        Map<String, List<PageInfo>> companyMap = new LinkedHashMap<>();

        for (SourceConfig src : sources) {
            PdfDocument doc = openDocs.get(src.path);
            int totalPages = doc.getNumberOfPages();

            for (int pageNum = 1; pageNum <= totalPages; pageNum++) {
                String pageText = PdfTextExtractor.getTextFromPage(
                        doc.getPage(pageNum),
                        new LocationTextExtractionStrategy()
                );

                CompanyAndType cat = src.extractor.extract(pageText);

                String rawCompany = cat.companyName;
                String companyKey = normalizeCompanyKey(rawCompany);
                String serviceType = normalizeServiceType(cat.serviceTypeRaw);

                BigDecimal amount = null;
                if (cat.totalAmountRaw != null) {
                    amount = parseAmount(cat.totalAmountRaw);
                }

                // 2nd file: accumulate service-type totals
                if (src.sortByService && amount != null) {
                    serviceTotalsByCompany
                            .computeIfAbsent(companyKey, k -> new LinkedHashMap<>())
                            .merge(serviceType, amount, BigDecimal::add);
                }

                // 1st file: accumulate CHSS total
                if (!src.sortByService && src.inputOrder == 1 && amount != null) {
                    chssTotalByCompany.merge(companyKey, amount, BigDecimal::add);
                }

                PageInfo info = new PageInfo(
                        src.path,
                        pageNum,
                        rawCompany,
                        companyKey,
                        serviceType,
                        src.namePriority,
                        src.inputOrder,
                        src.sortByService,
                        amount
                );

                companyMap
                        .computeIfAbsent(companyKey, k -> new ArrayList<>())
                        .add(info);
            }
        }

        // For each company, sort and create output PDF + Excel
        for (Map.Entry<String, List<PageInfo>> entry : companyMap.entrySet()) {
            String companyKey = entry.getKey();
            List<PageInfo> pages = entry.getValue();

            // sort: by inputOrder, then (if applicable) by service type, then by page number
            pages.sort((p1, p2) -> {
                if (p1.inputOrder != p2.inputOrder) {
                    return Integer.compare(p1.inputOrder, p2.inputOrder);
                }

                if (p1.sortByService && p2.sortByService) {
                    int o1 = serviceTypeOrder.getOrDefault(p1.serviceType, serviceTypeOrder.get("OTHERS"));
                    int o2 = serviceTypeOrder.getOrDefault(p2.serviceType, serviceTypeOrder.get("OTHERS"));
                    if (o1 != o2) {
                        return Integer.compare(o1, o2);
                    }
                }

                return Integer.compare(p1.pageNum, p2.pageNum);
            });

            // output company name: prefer lowest namePriority with a non-empty rawCompany
            String outputCompanyName = pages.stream()
                    .sorted(Comparator.comparingInt(p -> p.namePriority))
                    .map(p -> p.rawCompany)
                    .filter(Objects::nonNull)
                    .map(String::trim)
                    .filter(s -> !s.isEmpty())
                    .findFirst()
                    .orElse("UNKNOWN");

            String fileName = sanitizeForFilename(outputCompanyName);
            String pdfPath = outputDir + fileName + ".pdf";

            try (PdfDocument destDoc = new PdfDocument(new PdfWriter(pdfPath))) {
                for (PageInfo p : pages) {
                    PdfDocument srcDoc = openDocs.get(p.srcPath);
                    srcDoc.copyPagesTo(p.pageNum, p.pageNum, destDoc);
                }
            }

            System.out.println("PDF Created for " + fileName);

            Map<String, BigDecimal> totals = serviceTotalsByCompany.get(companyKey);
            BigDecimal chssAmount = chssTotalByCompany.get(companyKey);

            if ((totals != null && !totals.isEmpty()) || chssAmount != null) {
                generateExcelForCompany(outputDir, fileName, outputCompanyName, totals, chssAmount);
            }
        }

        for (PdfDocument doc : openDocs.values()) {
            doc.close();
        }
    }

    // ---------- helpers ----------

    /**
     * Normalized key for matching across PDFs:
     * remove spaces and dots, uppercase.
     */
    private String normalizeCompanyKey(String name) {
        if (name == null) return "UNKNOWN";
        String s = name.trim();
        if (s.isEmpty()) return "UNKNOWN";
        s = s.replaceAll("[ .]", "");  // remove spaces and dots
        s = s.toUpperCase();
        return s.isEmpty() ? "UNKNOWN" : s;
    }

    /**
     * Safe filename for output PDF/XLSX.
     */
    private String sanitizeForFilename(String name) {
        if (name == null || name.trim().isEmpty()) {
            return "UNKNOWN";
        }
        String s = name.replaceAll("[\\\\/:*?\"<>|]", "_").trim();
        return s.isEmpty() ? "UNKNOWN" : s;
    }

    /**
     * Map raw type into TAX / ACCOUNT / BPO / SECRETARY / OTHERS.
     */
    private String normalizeServiceType(String raw) {
        if (raw == null) {
            return "OTHERS";
        }
        String s = raw.trim().toUpperCase();

        if (s.startsWith("TAX")) {
            return "TAX";
        }
        if (s.contains("ACCOUNT")) {
            return "ACCOUNT";
        }
        if (s.contains("BPO")) {
            return "BPO";
        }
        if (s.contains("SECRET")) {
            return "SECRETARY";
        }
        return "OTHERS";
    }

    private BigDecimal parseAmount(String raw) {
        if (raw == null) return null;
        String s = raw.trim().replace(",", "");
        if (s.isEmpty()) return null;
        try {
            return new BigDecimal(s);
        } catch (NumberFormatException e) {
            return null;
        }
    }

    /**
     * Find the first number AFTER a marker in the text.
     */
    private String findAmountAfterMarker(String text, String marker) {
        if (text == null || marker == null) return null;
        int idx = text.indexOf(marker);
        if (idx < 0) return null;

        String after = text.substring(idx + marker.length());

        Pattern p = Pattern.compile("([0-9]{1,3}(?:,[0-9]{3})*(?:\\.[0-9]{2})?|[0-9]+(?:\\.[0-9]{2})?)");
        Matcher m = p.matcher(after);
        if (m.find()) {
            return m.group(1);
        }
        return null;
    }

    /**
     * Generate Excel for one company from Template.xlsx:
     * - A2: company name
     * - For col D codes TAX/ACC/BPO/SEC/OTHERS: put amount in col B
     * - For D = CHSS Invoice: put CHSS amount in B
     * - For row where A = Total: put grand total in B
     */
    private void generateExcelForCompany(String outputDir,
                                         String fileName,
                                         String displayCompanyName,
                                         Map<String, BigDecimal> totals,
                                         BigDecimal chssAmount) throws IOException {

        File tpl = new File(templatePath);
        if (!tpl.exists()) {
            System.out.println("Template not found: " + templatePath + ", skip Excel for " + displayCompanyName);
            return;
        }

        try (FileInputStream fis = new FileInputStream(tpl);
             Workbook wb = new XSSFWorkbook(fis)) {

            Sheet sheet = wb.getSheetAt(0); // Template has Sheet1 as first

            // A2: company name
            Row row2 = sheet.getRow(1);
            if (row2 == null) row2 = sheet.createRow(1);
            Cell companyCell = row2.getCell(0);
            if (companyCell == null) companyCell = row2.createCell(0);
            companyCell.setCellValue(displayCompanyName);

            // Build map of code in column D -> row index
            Map<String, Integer> codeRowIndex = new HashMap<>();
            for (Row row : sheet) {
                Cell codeCell = row.getCell(3); // column D
                if (codeCell != null && codeCell.getCellType() == CellType.STRING) {
                    String code = codeCell.getStringCellValue();
                    if (code != null) {
                        code = code.trim().toUpperCase();
                        if (!code.isEmpty()) {
                            codeRowIndex.put(code, row.getRowNum());
                        }
                    }
                }
            }

            // Fill service-type totals from 2nd file
            if (totals != null) {
                for (Map.Entry<String, BigDecimal> e : totals.entrySet()) {
                    String svcType = e.getKey();
                    BigDecimal amt = e.getValue();
                    if (amt == null) continue;

                    String code;
                    switch (svcType) {
                        case "TAX":
                            code = "TAX";
                            break;
                        case "ACCOUNT":
                            code = "ACC";
                            break;
                        case "BPO":
                            code = "BPO";
                            break;
                        case "SECRETARY":
                            code = "SEC";
                            break;
                        default:
                            code = "OTHERS";
                    }

                    Integer rowIdx = codeRowIndex.get(code);
                    if (rowIdx == null) continue;

                    Row row = sheet.getRow(rowIdx);
                    if (row == null) row = sheet.createRow(rowIdx);
                    Cell amountCell = row.getCell(1); // column B
                    if (amountCell == null) amountCell = row.createCell(1);
                    amountCell.setCellValue(amt.doubleValue());
                }
            }

            // CHSS amount from 1st file: map to row where D = "CHSS Invoice"
            if (chssAmount != null) {
                Integer chssRowIdx = codeRowIndex.get("CHSS INVOICE");
                if (chssRowIdx != null) {
                    Row chssRow = sheet.getRow(chssRowIdx);
                    if (chssRow == null) chssRow = sheet.createRow(chssRowIdx);
                    Cell chssCell = chssRow.getCell(1); // column B
                    if (chssCell == null) chssCell = chssRow.createCell(1);
                    chssCell.setCellValue(chssAmount.doubleValue());
                }
            }

            // Grand total = CHSS + all service-type totals
            BigDecimal grandTotal = BigDecimal.ZERO;
            if (chssAmount != null) {
                grandTotal = grandTotal.add(chssAmount);
            }
            if (totals != null) {
                for (BigDecimal amt : totals.values()) {
                    if (amt != null) {
                        grandTotal = grandTotal.add(amt);
                    }
                }
            }

            if (grandTotal.compareTo(BigDecimal.ZERO) > 0) {
                // Find row where A = "Total"
                Row totalRow = null;
                for (Row row : sheet) {
                    Cell labelCell = row.getCell(0); // column A
                    if (labelCell != null && labelCell.getCellType() == CellType.STRING) {
                        String text = labelCell.getStringCellValue();
                        if (text != null && text.trim().equalsIgnoreCase("Total")) {
                            totalRow = row;
                            break;
                        }
                    }
                }

                if (totalRow != null) {
                    Cell totalCell = totalRow.getCell(1); // column B
                    if (totalCell == null) totalCell = totalRow.createCell(1);
                    totalCell.setCellValue(grandTotal.doubleValue());
                } else {
                    System.out.println("No 'Total' row found in column A for " + displayCompanyName);
                }
            }

            String xlsxPath = outputDir + fileName + ".xlsx";
            try (FileOutputStream fos = new FileOutputStream(xlsxPath)) {
                wb.write(fos);
            }

            System.out.println("Excel Created for " + fileName);
        }
    }

    // ---------- inner classes ----------

    public class SourceConfig {
        public final String path;
        public final PageExtractor extractor;
        public final int namePriority;
        public final int inputOrder;
        public final boolean sortByService;

        public SourceConfig(String path,
                            PageExtractor extractor,
                            int namePriority,
                            int inputOrder,
                            boolean sortByService) {
            this.path = path;
            this.extractor = extractor;
            this.namePriority = namePriority;
            this.inputOrder = inputOrder;
            this.sortByService = sortByService;
        }
    }

    public class PageInfo {
        public final String srcPath;
        public final int pageNum;
        public final String rawCompany;
        public final String companyKey;
        public final String serviceType;
        public final int namePriority;
        public final int inputOrder;
        public final boolean sortByService;
        public final BigDecimal totalAmount;

        public PageInfo(String srcPath,
                        int pageNum,
                        String rawCompany,
                        String companyKey,
                        String serviceType,
                        int namePriority,
                        int inputOrder,
                        boolean sortByService,
                        BigDecimal totalAmount) {
            this.srcPath = srcPath;
            this.pageNum = pageNum;
            this.rawCompany = rawCompany;
            this.companyKey = companyKey;
            this.serviceType = serviceType;
            this.namePriority = namePriority;
            this.inputOrder = inputOrder;
            this.sortByService = sortByService;
            this.totalAmount = totalAmount;
        }
    }

    public class CompanyAndType {
        public final String companyName;
        public final String serviceTypeRaw;
        public final String totalAmountRaw;

        public CompanyAndType(String companyName, String serviceTypeRaw, String totalAmountRaw) {
            this.companyName = companyName;
            this.serviceTypeRaw = serviceTypeRaw;
            this.totalAmountRaw = totalAmountRaw;
        }
    }

    public interface PageExtractor {
        CompanyAndType extract(String pageText);
    }

    // ---------- 1st PDF extractor ----------
    // 1. Find "Invoice"
    // 2. Next line, parse "To : " -> company (until multi spaces)
    // 3. Search "Total payable inclusive of service tax :" and get amount after it
    public class InvoiceExtractor implements InvoiceSplitter.PageExtractor {

        @Override
        public InvoiceSplitter.CompanyAndType extract(String pageText) {
            if (pageText == null) {
                return new InvoiceSplitter.CompanyAndType(null, null, null);
            }

            String[] lines = pageText.split("\\r?\\n");
            String targetLine = null;

            for (int i = 0; i < lines.length; i++) {
                if (lines[i].contains("Invoice")) {
                    if (i + 1 < lines.length) {
                        targetLine = lines[i + 1];
                    }
                    break;
                }
            }

            if (targetLine == null) {
                for (String line : lines) {
                    if (line.contains("To : ")) {
                        targetLine = line;
                        break;
                    }
                }
            }

            String company = null;
            if (targetLine != null) {
                String marker = "To  : ";
                int idx = targetLine.indexOf(marker);
                if (idx < 0) {
                    marker = "To : ";
                    idx = targetLine.indexOf(marker);
                }
                if (idx >= 0) {
                    String after = targetLine.substring(idx + marker.length());
                    int parenIdx = after.indexOf(" Doc");
                    if (parenIdx >= 0) {
                        company = after.substring(0, parenIdx).trim();
                    } else {
                        company = after.trim();
                    }
                }
            }

            String totalAmount = findAmountAfterMarker(
                    pageText,
                    "Total payable inclusive of service tax :"
            );

            return new InvoiceSplitter.CompanyAndType(company, null, totalAmount);
        }
    }

    // ---------- 2nd PDF extractor ----------
    // 1. Find "Invoice"
    // 2. Next line, parse:
    //    - "To  : " or "To : " -> company (until " (")
    //    - "Service Type :" -> type (until multi spaces)
    // 3. Search "Total :" and get amount after it
    public class InvoiceByTypeExtractor implements PageExtractor {

        @Override
        public CompanyAndType extract(String pageText) {
            if (pageText == null) {
                return new CompanyAndType(null, null, null);
            }

            String[] lines = pageText.split("\\r?\\n");
            String targetLine = null;

            for (int i = 0; i < lines.length; i++) {
                if (lines[i].contains("Invoice")) {
                    if (i + 1 < lines.length) {
                        targetLine = lines[i + 1];
                    }
                    break;
                }
            }

            if (targetLine == null) {
                for (String line : lines) {
                    if (line.contains("To  : ") || line.contains("To : ")) {
                        targetLine = line;
                        break;
                    }
                }
            }

            String company = null;
            String serviceType = null;

            if (targetLine != null) {
                // company
                String marker = "To  : ";
                int idx = targetLine.indexOf(marker);
                if (idx < 0) {
                    marker = "To : ";
                    idx = targetLine.indexOf(marker);
                }
                if (idx >= 0) {
                    String after = targetLine.substring(idx + marker.length());
                    int parenIdx = after.indexOf(" (");
                    if (parenIdx >= 0) {
                        company = after.substring(0, parenIdx).trim();
                    } else {
                        company = after.trim();
                    }
                }

                // service type
                String stMarker = "Service Type :";
                int stIdx = targetLine.indexOf(stMarker);
                if (stIdx >= 0) {
                    String after = targetLine.substring(stIdx + stMarker.length()).trim();
                    if (after.startsWith("\"")) {
                        after = after.substring(1).trim();
                    }
                    String[] stParts = after.split("\\s{2,}");
                    String raw = (stParts.length > 0 ? stParts[0] : after).replace("\"", "").trim();
                    if (!raw.isEmpty()) {
                        serviceType = raw;
                    }
                }
            }

            String totalAmount = findAmountAfterMarker(pageText, "Total :");

            return new CompanyAndType(company, serviceType, totalAmount);
        }
    }

    // ---------- 3rd PDF extractor ----------
    // 1. Find "Customer"
    // 2. Next line, company = from first char to multi spaces
    public class StatementOfAccountExtractor implements PageExtractor {

        @Override
        public CompanyAndType extract(String pageText) {
            if (pageText == null) {
                return new CompanyAndType(null, null, null);
            }

            String[] lines = pageText.split("\\r?\\n");
            String targetLine = null;

            for (int i = 0; i < lines.length; i++) {
                if (lines[i].contains("Statement of Account")) {
                    if (i + 1 < lines.length) {
                        targetLine = lines[i + 1];
                    }
                    break;
                }
            }

            String company = null;
            if (targetLine != null) {
                String[] parts = targetLine.split("\\s{2,}");
                company = (parts.length > 0 ? parts[0] : targetLine).trim();
            }

            return new CompanyAndType(company, null, null);
        }
    }
}


