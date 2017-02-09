/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pl.stepinski.restartnumberingintables_docx4j;

import java.io.File;
import java.math.BigInteger;
import javax.xml.bind.JAXBElement;
import org.docx4j.XmlUtils;
import org.docx4j.jaxb.Context;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.WordprocessingML.NumberingDefinitionsPart;
import org.docx4j.wml.Numbering;
import org.docx4j.wml.Tbl;
import org.docx4j.wml.Tc;
import org.docx4j.wml.Tr;
import org.docx4j.wml.P;
import org.docx4j.wml.PPrBase;

/**
 *
 * @author piotr
 */
public class App {
    
    static org.docx4j.wml.ObjectFactory factory = Context.getWmlObjectFactory(); 
    
    public static void main(String[] args){
        try {
            File template = new File("template.docx");
        
            WordprocessingMLPackage document = WordprocessingMLPackage.load(template);
            
            MainDocumentPart mdp = document.getMainDocumentPart();
            
            P p1 = (P) XmlUtils.unwrap(mdp.getJaxbElement().getBody().getEGBlockLevelElts().get(0));
            P p2 = (P) XmlUtils.unwrap(mdp.getJaxbElement().getBody().getEGBlockLevelElts().get(1));

            Tbl table = (Tbl) XmlUtils.unwrap(mdp.getJaxbElement().getBody().getEGBlockLevelElts().get(2));
            Tr tableRow = (Tr) table.getContent().get(1);
            
            NumberingDefinitionsPart ndp = new NumberingDefinitionsPart();
            mdp.addTargetPart(ndp);
            ndp.setJaxbElement( (Numbering) XmlUtils.unmarshalString(initialNumbering) );

            for( int r = 0; r < 5; r++ ){
                P tableName = XmlUtils.deepCopy(p1);
                P tableComent = XmlUtils.deepCopy(p2);
                Tbl tempTable = XmlUtils.deepCopy(table);
                
                // Ok, lets restart numbering
                long newNumId = ndp.restart(1, 0, 1);
                
                for(int i=0; i<3; i++){
                    Tr row = XmlUtils.deepCopy(tableRow);
                    P p = createNumberedParagraph(newNumId, 0, "" );
                    
                    JAXBElement<Tc> firstCell = (JAXBElement<Tc>) row.getContent().get(0);
                    firstCell.getValue().getContent().set(0, p);
                    
                    tempTable.getContent().add(row);
                }
                
                mdp.addObject(tableName);
                mdp.addObject(tableComent);
                
                
                // remove second row comming from template table
                tempTable.getContent().remove(1);
                
                mdp.addObject(tempTable);
            }
            
            // remove first 3 elements comming from template file
            mdp.getContent().remove(0);
            mdp.getContent().remove(0);
            mdp.getContent().remove(0);
            
            document.save(new File("output.docx"));
            System.out.println("DONE");
        } catch (Exception ex) {
            System.out.println("ERROR: " + ex.toString());
        }
    }
    
    private static P createNumberedParagraph(long numId, long ilvl, String paragraphText ) {
		
		P  p = factory.createP();

		org.docx4j.wml.Text  t = factory.createText();
		t.setValue(paragraphText);

		org.docx4j.wml.R  run = factory.createR();
		run.getContent().add(t);		
		
		p.getContent().add(run);
						
	    org.docx4j.wml.PPr ppr = factory.createPPr();	    
	    p.setPPr( ppr );
	    
	    // Create and add <w:numPr>
	    PPrBase.NumPr numPr =  factory.createPPrBaseNumPr();
	    ppr.setNumPr(numPr);
	    
	    // The <w:ilvl> element
	    PPrBase.NumPr.Ilvl ilvlElement = factory.createPPrBaseNumPrIlvl();
	    numPr.setIlvl(ilvlElement);
	    ilvlElement.setVal(BigInteger.valueOf(ilvl));
	    	    
	    // The <w:numId> element
	    PPrBase.NumPr.NumId numIdElement = factory.createPPrBaseNumPrNumId();
	    numPr.setNumId(numIdElement);
	    numIdElement.setVal(BigInteger.valueOf(numId));
	    
		return p;
		
	}
	static final String initialNumbering = "<w:numbering xmlns:ve=\"http://schemas.openxmlformats.org/markup-compatibility/2006\" xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\" xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\" xmlns:v=\"urn:schemas-microsoft-com:vml\" xmlns:wp=\"http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing\" xmlns:w10=\"urn:schemas-microsoft-com:office:word\" xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\" xmlns:wne=\"http://schemas.microsoft.com/office/word/2006/wordml\">"
			    + "<w:abstractNum w:abstractNumId=\"0\">"
			    + "<w:nsid w:val=\"2DD860C0\"/>"
			    + "<w:multiLevelType w:val=\"multilevel\"/>"
			    + "<w:tmpl w:val=\"0409001D\"/>"
			    + "<w:lvl w:ilvl=\"0\">"
			        + "<w:start w:val=\"1\"/>"
			        + "<w:numFmt w:val=\"decimal\"/>"
			        + "<w:lvlText w:val=\"%1)\"/>"
			        + "<w:lvlJc w:val=\"left\"/>"
			        + "<w:pPr>"
			            + "<w:ind w:left=\"360\" w:hanging=\"360\"/>"
			        + "</w:pPr>"
			    + "</w:lvl>"
			    + "<w:lvl w:ilvl=\"1\">"
			        + "<w:start w:val=\"1\"/>"
			        + "<w:numFmt w:val=\"lowerLetter\"/>"
			        + "<w:lvlText w:val=\"%2)\"/>"
			        + "<w:lvlJc w:val=\"left\"/>"
			        + "<w:pPr>"
			            + "<w:ind w:left=\"720\" w:hanging=\"360\"/>"
			        + "</w:pPr>"
			    + "</w:lvl>"
			    + "<w:lvl w:ilvl=\"2\">"
			        + "<w:start w:val=\"1\"/>"
			        + "<w:numFmt w:val=\"lowerRoman\"/>"
			        + "<w:lvlText w:val=\"%3)\"/>"
			        + "<w:lvlJc w:val=\"left\"/>"
			        + "<w:pPr>"
			            + "<w:ind w:left=\"1080\" w:hanging=\"360\"/>"
			        + "</w:pPr>"
			    + "</w:lvl>"
			    + "<w:lvl w:ilvl=\"3\">"
			        + "<w:start w:val=\"1\"/>"
			        + "<w:numFmt w:val=\"decimal\"/>"
			        + "<w:lvlText w:val=\"(%4)\"/>"
			        + "<w:lvlJc w:val=\"left\"/>"
			        + "<w:pPr>"
			            + "<w:ind w:left=\"1440\" w:hanging=\"360\"/>"
			        + "</w:pPr>"
			    + "</w:lvl>"
			    + "<w:lvl w:ilvl=\"4\">"
			        + "<w:start w:val=\"1\"/>"
			        + "<w:numFmt w:val=\"lowerLetter\"/>"
			        + "<w:lvlText w:val=\"(%5)\"/>"
			        + "<w:lvlJc w:val=\"left\"/>"
			        + "<w:pPr>"
			            + "<w:ind w:left=\"1800\" w:hanging=\"360\"/>"
			        + "</w:pPr>"
			    + "</w:lvl>"
			    + "<w:lvl w:ilvl=\"5\">"
			        + "<w:start w:val=\"1\"/>"
			        + "<w:numFmt w:val=\"lowerRoman\"/>"
			        + "<w:lvlText w:val=\"(%6)\"/>"
			        + "<w:lvlJc w:val=\"left\"/>"
			        + "<w:pPr>"
			            + "<w:ind w:left=\"2160\" w:hanging=\"360\"/>"
			        + "</w:pPr>"
			    + "</w:lvl>"
			    + "<w:lvl w:ilvl=\"6\">"
			        + "<w:start w:val=\"1\"/>"
			        + "<w:numFmt w:val=\"decimal\"/>"
			        + "<w:lvlText w:val=\"%7.\"/>"
			        + "<w:lvlJc w:val=\"left\"/>"
			        + "<w:pPr>"
			            + "<w:ind w:left=\"2520\" w:hanging=\"360\"/>"
			        + "</w:pPr>"
			    + "</w:lvl>"
			    + "<w:lvl w:ilvl=\"7\">"
			        + "<w:start w:val=\"1\"/>"
			        + "<w:numFmt w:val=\"lowerLetter\"/>"
			        + "<w:lvlText w:val=\"%8.\"/>"
			        + "<w:lvlJc w:val=\"left\"/>"
			        + "<w:pPr>"
			            + "<w:ind w:left=\"2880\" w:hanging=\"360\"/>"
			        + "</w:pPr>"
			    + "</w:lvl>"
			    + "<w:lvl w:ilvl=\"8\">"
			        + "<w:start w:val=\"1\"/>"
			        + "<w:numFmt w:val=\"lowerRoman\"/>"
			        + "<w:lvlText w:val=\"%9.\"/>"
			        + "<w:lvlJc w:val=\"left\"/>"
			        + "<w:pPr>"
			            + "<w:ind w:left=\"3240\" w:hanging=\"360\"/>"
			        + "</w:pPr>"
			    + "</w:lvl>"
			+ "</w:abstractNum>"
			+ "<w:num w:numId=\"1\">"
			    + "<w:abstractNumId w:val=\"0\"/>"
			 + "</w:num>"
			+ "</w:numbering>";
}
