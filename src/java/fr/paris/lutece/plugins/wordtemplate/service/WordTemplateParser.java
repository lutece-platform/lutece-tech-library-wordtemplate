/*
 * Copyright (c) 2002-2020, City of Paris
 * All rights reserved.
 *
 * Redistribution and use in source and binary forms, with or without
 * modification, are permitted provided that the following conditions
 * are met:
 *
 *  1. Redistributions of source code must retain the above copyright notice
 *     and the following disclaimer.
 *
 *  2. Redistributions in binary form must reproduce the above copyright notice
 *     and the following disclaimer in the documentation and/or other materials
 *     provided with the distribution.
 *
 *  3. Neither the name of 'Mairie de Paris' nor 'Lutece' nor the names of its
 *     contributors may be used to endorse or promote products derived from
 *     this software without specific prior written permission.
 *
 * THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
 * AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
 * IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
 * ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDERS OR CONTRIBUTORS BE
 * LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
 * CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
 * SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
 * INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
 * CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
 * ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
 * POSSIBILITY OF SUCH DAMAGE.
 *
 * License 1.0
 */
package fr.paris.lutece.plugins.wordtemplate.service;

import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import fr.paris.lutece.plugins.wordtemplate.business.IWordTemplateElement;
import fr.paris.lutece.plugins.wordtemplate.business.WordTemplate;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFFooter;
import org.apache.poi.xwpf.usermodel.XWPFHeader;

/**
 * Parser of word templates
 */
public class WordTemplateParser
{
    private static final String INSTRUCTION_PATTERN = "\\$\\{.*?\\}|</?#.*?>";

    /**
     * Replace properties by their value
     *
     * @param document
     * @param model
     * @return 
     * @throws Exception
     */
    public WordTemplate parse( XWPFDocument document, Map<String, Object> model ) throws Exception
    {
        WordTemplate wordTemplate = new WordTemplate( );
        List<IWordTemplateElement> listTemplateElements = new ArrayList<>( );

        // Parse headers
        for ( XWPFHeader header : document.getHeaderList( ) )
        {
            listTemplateElements.addAll( findInstr( header ) );
        }
        // Parse footers
        for ( XWPFFooter footer : document.getFooterList( ) )
        {
            listTemplateElements.addAll( findInstr( footer ) );
        }

        // Parse the document
        listTemplateElements.addAll( findInstr( document ) );

        wordTemplate.setListInstructions( listTemplateElements );

        return wordTemplate;
    }

    /**
     *
     * @param body
     */
    public static void visitBody( IBody body )
    {
        for ( IBodyElement bodyElement : body.getBodyElements( ) )
        {
            if ( bodyElement.getElementType( ).equals( BodyElementType.PARAGRAPH ) )
            {
                for ( XWPFRun run : ( (XWPFParagraph) bodyElement ).getRuns( ) )
                {
                    //action
                }
            }
            if ( bodyElement.getElementType( ).equals( BodyElementType.TABLE ) )
            {
                for ( XWPFTableRow row : ( (XWPFTable) bodyElement ).getRows( ) )
                {
                    for ( XWPFTableCell cell : row.getTableCells( ) )
                    {
                        visitBody( cell );
                    }
                }
            }
        }
    }

    /**
     *
     * @param body
     * @return
     */
    private List<IWordTemplateElement> findInstr( IBody body )
    {
        List<IWordTemplateElement> listInstruction = new ArrayList<>( );

        for ( IBodyElement bodyElement : body.getBodyElements( ) )
        {
            if ( bodyElement.getElementType( ).equals( BodyElementType.PARAGRAPH ) )
            {
                listInstruction.addAll( findInstr( (XWPFParagraph) bodyElement ) );
            }

            if ( bodyElement.getElementType( ).equals( BodyElementType.TABLE ) )
            {
                for ( XWPFTableRow row : ( (XWPFTable) bodyElement ).getRows( ) )
                {
                    for ( XWPFTableCell cell : row.getTableCells( ) )
                    {
                        findInstr( cell );
                    }
                }
            }
        }

        return listInstruction;
    }

    /**
     *
     * @param paragraph
     * @return
     */
    private List<IWordTemplateElement> findInstr( XWPFParagraph paragraph )
    {
        List<IWordTemplateElement> listInstruction = new ArrayList<>( );
        String text = paragraph.getParagraphText( );
        Pattern pattern = Pattern.compile( INSTRUCTION_PATTERN );
        Matcher matcher = pattern.matcher( text );

        while ( matcher.find( ) )
        {
            int start = matcher.start( );
            int end = matcher.end( ) - 1;
            formatInstr( paragraph, start, end );
            for ( XWPFRun run : paragraph.getRuns( ) )
            {
                if ( run.toString( ).equals( matcher.group( ) ) )
                {
                    InstructionService instructionService = InstructionService.init( );
                    IWordTemplateElement element = instructionService.createInstruction( run.toString( ), run );
                    listInstruction.add( element );
                    break;
                }
            }
        }

        return listInstruction;
    }

    /**
     *
     * @param paragraph
     * @param start
     * @param end
     */
    private void formatInstr( XWPFParagraph paragraph, int start, int end )
    {
        boolean startFound = false;
        boolean endFound = false;

        while ( !endFound )
        {
            int pos = 0, numRun = -1, nextPos;

            for ( XWPFRun run : paragraph.getRuns( ) )
            {
                numRun++;
                nextPos = pos + run.toString( ).length( );

                if ( start < pos && end >= nextPos )
                {
                    XWPFRun prevRun = paragraph.getRuns( ).get( numRun - 1 );
                    prevRun.setText( prevRun.toString( ) + run.toString( ), 0 );
                    paragraph.removeRun( numRun );
                    break;
                }
                if ( start >= pos && start < nextPos && !startFound )
                {
                    int startPosInRun = start - pos;
                    startFound = true;
                    if ( start > pos )
                    {
                        WordService.splitRun( run, startPosInRun );
                        break;
                    }
                }
                if ( end >= pos && end < nextPos )
                {
                    int endPosInRun = end - pos;
                    endFound = true;

                    WordService.splitRun( run, endPosInRun + 1 );

                    if ( start < pos )
                    {
                        XWPFRun prevRun = paragraph.getRuns( ).get( numRun - 1 );
                        run = paragraph.getRuns( ).get( numRun );
                        prevRun.setText( prevRun.toString( ) + run.toString( ), 0 );
                        paragraph.removeRun( numRun );
                    }
                    break;
                }

                pos = nextPos;
            }
        }
    }
}
