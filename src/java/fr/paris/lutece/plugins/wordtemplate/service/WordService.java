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

import fr.paris.lutece.plugins.wordtemplate.exception.WordTemplateException;
import java.util.ArrayList;
import java.util.List;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBody;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.apache.xmlbeans.XmlCursor;
import org.apache.xmlbeans.XmlException;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTPPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRow;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTblPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTc;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTcPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTTrPr;

/**
 * Service for manipulation of objects of a word document
 */
public class WordService
{

    /***
     * Get the content of body
     * 
     * @param body
     * @return the content of the body
     * @throws XmlException
     */
    public String getContent( IBody body ) throws XmlException
    {
        StringBuilder stringBuilder = new StringBuilder( );

        for ( IBodyElement bodyElement : body.getBodyElements( ) )
        {
            if ( bodyElement.getElementType( ).equals( BodyElementType.PARAGRAPH ) )
            {
                for ( XWPFRun run : ( (XWPFParagraph) bodyElement ).getRuns( ) )
                {
                    String text = run.getText( 0 );
                    stringBuilder.append( text );
                }
            }
            if ( bodyElement.getElementType( ).equals( BodyElementType.TABLE ) )
            {
                for ( XWPFTableRow row : ( (XWPFTable) bodyElement ).getRows( ) )
                {
                    for ( XWPFTableCell cell : row.getTableCells( ) )
                    {
                        getContent( cell );
                    }
                }
            }
        }

        return stringBuilder.toString( );
    }

    /**
     * Clone an IBodyElement
     * 
     * @param clone
     *            the cloned IBodyElement
     * @param source
     *            the source for IBodyElement
     */
    public static void cloneBodyElement( IBodyElement clone, IBodyElement source )
    {
        if ( clone.getElementType( ).equals( BodyElementType.PARAGRAPH ) && source.getElementType( ).equals( BodyElementType.PARAGRAPH ) )
        {
            cloneParagraph( (XWPFParagraph) clone, (XWPFParagraph) source, false );
            return;
        }
        if ( clone.getElementType( ).equals( BodyElementType.TABLE ) && source.getElementType( ).equals( BodyElementType.TABLE ) )
        {
            cloneTable( (XWPFTable) clone, (XWPFTable) source, false );
            return;
        }
        throw new WordTemplateException( "Try to clone one BodyElement into another but with different type" );
    }

    /**
     * Clone a paragraph
     * 
     * @param clone
     *            the cloned paragraph
     * @param source
     *            the source for paragraph
     * @param isEmpty
     */
    public static void cloneParagraph( XWPFParagraph clone, XWPFParagraph source, boolean isEmpty )
    {
        CTPPr pPr = clone.getCTP( ).isSetPPr( ) ? clone.getCTP( ).getPPr( ) : clone.getCTP( ).addNewPPr( );
        pPr.set( source.getCTP( ).getPPr( ) );

        if ( isEmpty )
        {
            return;
        }

        for ( XWPFRun run : source.getRuns( ) )
        {
            XWPFRun newRun = clone.createRun( );
            cloneRun( newRun, run, false );
        }
    }

    /**
     * Clone a run
     * 
     * @param clone
     *            the cloned run
     * @param source
     *            the source for run
     * @param isEmpty
     */
    public static void cloneRun( XWPFRun clone, XWPFRun source, boolean isEmpty )
    {
        CTRPr rPr = clone.getCTR( ).isSetRPr( ) ? clone.getCTR( ).getRPr( ) : clone.getCTR( ).addNewRPr( );
        rPr.set( source.getCTR( ).getRPr( ) );

        if ( isEmpty )
        {
            return;
        }

        clone.setText( source.getText( 0 ) );
    }

    /**
     * Clone a table
     * 
     * @param clone
     *            the cloned table
     * @param source
     *            the source for table
     * @param isEmpty
     */
    public static void cloneTable( XWPFTable clone, XWPFTable source, boolean isEmpty )
    {
        CTTblPr tblPr = clone.getCTTbl( ).getTblPr( ) != null ? clone.getCTTbl( ).getTblPr( ) : clone.getCTTbl( ).addNewTblPr( );
        tblPr.set( source.getCTTbl( ).getTblPr( ) );

        if ( isEmpty )
        {
            return;
        }

        boolean first = true;
        List<XWPFTableRow [ ]> newRows = new ArrayList<>( );

        for ( XWPFTableRow row : source.getRows( ) )
        {
            XWPFTableRow newRow;
            if ( first && clone.getRow( 0 ) != null )
            {
                newRow = clone.getRow( 0 );
            }
            else
            {
                newRow = clone.createRow( );
            }
            XWPFTableRow [ ] twoRows = {
                    newRow, row
            };
            newRows.add( twoRows );
            first = false;
        }

        for ( XWPFTableRow [ ] newRow : newRows )
        {
            cloneTableRow( newRow [0], newRow [1], false );
        }
    }

    /**
     * Clone a table row
     * 
     * @param clone
     *            the cloned table row
     * @param source
     *            the source for table row
     * @param isEmpty
     */
    public static void cloneTableRow( XWPFTableRow clone, XWPFTableRow source, boolean isEmpty )
    {
        CTTrPr trPr = clone.getCtRow( ).getTrPr( ) != null ? clone.getCtRow( ).getTrPr( ) : clone.getCtRow( ).addNewTrPr( );
        trPr.set( source.getCtRow( ).getTrPr( ) );

        if ( isEmpty )
        {
            return;
        }

        boolean first = true;

        for ( XWPFTableCell cell : source.getTableCells( ) )
        {
            XWPFTableCell newCell;
            if ( first && clone.getCell( 0 ) != null )
            {
                newCell = clone.getCell( 0 );
            }
            else
            {
                newCell = clone.createCell( );
            }
            cloneTableCell( newCell, cell );
            first = false;
        }
    }

    /**
     * Clone a table celle
     * 
     * @param clone
     *            the cloned table celle
     * @param source
     *            the source for table celle
     */
    public static void cloneTableCell( XWPFTableCell clone, XWPFTableCell source )
    {
        cloneTableCell( clone, source, 0, source.getBodyElements( ).size( ) );
    }

    /**
     * Clone a table celle
     * 
     * @param clone
     *            the cloned table celle
     * @param source
     *            the source for table celle
     * @param fromIndex
     * @param toIndex
     */
    public static void cloneTableCell( XWPFTableCell clone, XWPFTableCell source, int fromIndex, int toIndex )
    {
        CTTcPr tcPr = clone.getCTTc( ).getTcPr( ) != null ? clone.getCTTc( ).getTcPr( ) : clone.getCTTc( ).addNewTcPr( );
        tcPr.set( source.getCTTc( ).getTcPr( ) );

        XWPFParagraph firstParagraph = clone.getParagraphs( ).get( 0 );
        XmlCursor cursor = firstParagraph.getCTP( ).newCursor( );

        if ( !( fromIndex >= 0 && fromIndex <= toIndex && toIndex <= source.getBodyElements( ).size( ) ) )
        {
            return;
        }

        for ( int i = fromIndex; i < toIndex; i++ )
        {
            IBodyElement bodyElement = source.getBodyElements( ).get( i );
            if ( bodyElement.getElementType( ).equals( BodyElementType.PARAGRAPH ) )
            {
                XWPFParagraph newParagraph = clone.insertNewParagraph( cursor );
                cloneParagraph( newParagraph, (XWPFParagraph) bodyElement, false );
                cursor.dispose( );
                cursor = newParagraph.getCTP( ).newCursor( );
                cursor.toNextSibling( );
            }
            if ( bodyElement.getElementType( ).equals( BodyElementType.TABLE ) )
            {
                XWPFTable newTable = clone.insertNewTbl( cursor );
                cloneTable( newTable, (XWPFTable) bodyElement, false );
                cursor.dispose( );
                cursor = newTable.getCTTbl( ).newCursor( );
                cursor.toNextSibling( );
            }
        }
        cursor.dispose( );
        clone.removeParagraph( clone.getParagraphs( ).size( ) - 1 );
    }

    /**
     * Insert a table
     * 
     * @param body
     * @param table
     *            the table to insert
     * @param posDest
     * @return
     */
    public static XWPFTable insertTable( IBody body, XWPFTable table, int posDest )
    {
        IBodyElement bodyElement;

        if ( posDest > body.getBodyElements( ).size( ) || posDest < 0 )
        {
            return null;
        }

        if ( posDest == body.getBodyElements( ).size( ) )
        {
            bodyElement = body.getBodyElements( ).get( posDest - 1 );
        }
        else
        {
            bodyElement = body.getBodyElements( ).get( posDest );
        }

        XmlCursor cursor = getCursor( bodyElement );

        if ( posDest == body.getBodyElements( ).size( ) )
        {
            cursor.toParent( );
            cursor.toEndToken( );
        }

        XWPFTable newTable = insertTable( body, table, cursor );
        return newTable;
    }

    /**
     * Insert a table
     * 
     * @param body
     * @param table
     *            the table to insert
     * @param cursor
     * @return
     */
    public static XWPFTable insertTable( IBody body, XWPFTable table, XmlCursor cursor )
    {
        XWPFTable newTable = body.insertNewTbl( cursor );
        WordService.cloneTable( newTable, table, false );
        return newTable;
    }

    /**
     * Insert a table row
     * 
     * @param table
     * @param tableRow
     * @param posDest
     * @return
     */
    public static XWPFTableRow insertTableRow( XWPFTable table, XWPFTableRow tableRow, int posDest )
    {
        table.getCTTbl( ).insertNewTr( posDest );
        table.getCTTbl( ).setTrArray( posDest, tableRow.getCtRow( ) );
        XWPFTableRow newTableRow = new XWPFTableRow( table.getCTTbl( ).getTrArray( posDest ), table );

        table.getRows( ).add( posDest, newTableRow );

        return newTableRow;
    }

    /**
     * Insert a table row
     * 
     * @param table
     * @param tableRow
     * @param cursor
     * @return
     */
    public static XWPFTableRow insertTableRow( XWPFTable table, XWPFTableRow tableRow, XmlCursor cursor )
    {
        XWPFTableRow cursorTableRow = table.getRow( (CTRow) cursor.getObject( ) );
        int posDest = table.getRows( ).indexOf( cursorTableRow );
        return insertTableRow( table, tableRow, posDest );
    }

    /**
     * Insert a table cell
     * 
     * @param tableRow
     * @param tableCell
     * @param posDest
     * @return
     */
    public static XWPFTableCell insertTableCell( XWPFTableRow tableRow, XWPFTableCell tableCell, int posDest )
    {
        tableCell = addTableCell( tableRow, posDest );
        cloneTableCell( tableRow.getCell( posDest ), tableCell );
        return tableRow.getCell( posDest );
    }

    /**
     * Insert a table cell
     * 
     * @param tableRow
     * @param tableCell
     * @param cursor
     * @return
     */
    public static XWPFTableCell insertTableCell( XWPFTableRow tableRow, XWPFTableCell tableCell, XmlCursor cursor )
    {
        XWPFTableCell cursorTableCell = tableRow.getTableCell( (CTTc) cursor.getObject( ) );
        int posDest = tableRow.getTableCells( ).indexOf( cursorTableCell );
        return insertTableCell( tableRow, tableCell, posDest );
    }

    /**
     * Insert a paragraph
     * 
     * @param body
     * @param paragraph
     * @param posDest
     * @return
     */
    public static XWPFParagraph insertParagraph( IBody body, XWPFParagraph paragraph, int posDest )
    {
        IBodyElement bodyElement;

        if ( posDest > body.getBodyElements( ).size( ) || posDest < 0 )
        {
            return null;
        }

        if ( posDest == body.getBodyElements( ).size( ) )
        {
            bodyElement = body.getBodyElements( ).get( posDest - 1 );
        }
        else
        {
            bodyElement = body.getBodyElements( ).get( posDest );
        }

        XmlCursor cursor = getCursor( bodyElement );

        if ( posDest == body.getBodyElements( ).size( ) )
        {
            cursor.toParent( );
            cursor.toEndToken( );
        }

        return insertParagraph( body, paragraph, cursor );
    }

    /**
     * Insert a paragraph
     * 
     * @param body
     * @param paragraph
     * @param cursor
     * @return
     */
    public static XWPFParagraph insertParagraph( IBody body, XWPFParagraph paragraph, XmlCursor cursor )
    {
        XWPFParagraph newParagraph = body.insertNewParagraph( cursor );
        WordService.cloneParagraph( newParagraph, paragraph, false );
        return newParagraph;
    }

    /**
     * Insert a run
     * 
     * @param paragraphe
     * @param run
     * @param posDest
     * @return
     */
    public static XWPFRun insertRun( XWPFParagraph paragraphe, XWPFRun run, int posDest )
    {
        XWPFRun newRun = paragraphe.insertNewRun( posDest );
        cloneRun( newRun, run, false );
        return newRun;
    }

    /**
     * Insert a run
     * 
     * @param paragraphe
     * @param run
     * @param cursor
     * @return
     */
    public static XWPFRun insertRun( XWPFParagraph paragraphe, XWPFRun run, XmlCursor cursor )
    {
        XWPFRun cursorRun = paragraphe.getRun( (CTR) cursor.getObject( ) );
        int posDest = paragraphe.getRuns( ).indexOf( cursorRun );
        return insertRun( paragraphe, run, posDest );
    }

    /**
     * Split a paragraph
     *
     * @param paragraph
     * @param pos
     */
    public static void splitParagraph( XWPFParagraph paragraph, int pos )
    {
        if ( !( pos > 0 && pos < paragraph.getRuns( ).size( ) ) )
        {
            return;
        }
        XWPFParagraph beforeParagraph = WordService.insertParagraph( paragraph.getBody( ), paragraph, paragraph.getCTP( ).newCursor( ) );
        while ( beforeParagraph.removeRun( pos ) )
        {
        }
        while ( pos > 0 )
        {
            pos--;
            paragraph.removeRun( pos );
        }
    }

    /**
     * Split a run
     *
     * @param run
     * @param pos
     */
    public static void splitRun( XWPFRun run, int pos )
    {
        if ( !( run.getParent( ) instanceof XWPFParagraph ) )
        {
            return;
        }
        if ( !( pos > 0 && pos < run.toString( ).length( ) ) )
        {
            return;
        }
        String beforeText = run.toString( ).substring( 0, pos );
        String afterText = run.toString( ).substring( pos );
        XWPFParagraph paragraph = (XWPFParagraph) run.getParent( );
        int numRun = paragraph.getRuns( ).indexOf( run );
        XWPFRun beforeRun = paragraph.insertNewRun( numRun );
        WordService.cloneRun( beforeRun, run, true );
        beforeRun.setText( beforeText, 0 );
        run.setText( afterText, 0 );
    }

    /**
     * Split a table
     *
     * @param table
     * @param pos
     */
    public static void splitTable( XWPFTable table, int pos )
    {
        if ( !( pos > 0 && pos < table.getRows( ).size( ) ) )
        {
            return;
        }
        XWPFTable beforeTable = WordService.insertTable( table.getBody( ), table, table.getCTTbl( ).newCursor( ) );
        while ( beforeTable.removeRow( pos ) )
        {
        }
        while ( pos > 0 )
        {
            pos--;
            table.removeRow( pos );
        }
    }

    /**
     * Split a table row
     *
     * @param tableRow
     * @param pos
     */
    public static void splitTableRow( XWPFTableRow tableRow, int pos )
    {
        if ( !( pos > 0 && pos < tableRow.getTableCells( ).size( ) ) )
        {
            return;
        }
        XWPFTableRow beforeTableRow = WordService.insertTableRow( tableRow.getTable( ), tableRow, tableRow.getCtRow( ).newCursor( ) );
        while ( WordService.removeTableCell( beforeTableRow, pos ) )
        {
        }
        while ( pos > 0 )
        {
            pos--;
            WordService.removeTableCell( tableRow, pos );
        }
    }

    /**
     * Split a table cell
     *
     * @param tableCell
     * @param pos
     */
    public static void splitTableCell( XWPFTableCell tableCell, int pos )
    {
        if ( !( pos > 0 && pos < tableCell.getBodyElements( ).size( ) ) )
        {
            return;
        }
        XWPFTableRow tableRow = tableCell.getTableRow( );
        int posCell = tableRow.getTableCells( ).indexOf( tableCell );
        XWPFTableCell beforeTableCell1 = WordService.addTableCell( tableRow, posCell );
        XWPFTableCell beforeTableCell2 = WordService.addTableCell( tableRow, posCell );
        WordService.cloneTableCell( beforeTableCell2, tableCell, 0, pos );
        WordService.cloneTableCell( beforeTableCell1, tableCell, pos, tableCell.getBodyElements( ).size( ) );
        WordService.removeTableCell( tableRow, posCell + 2 );
    }

    /**
     * Add a table cell
     * 
     * @param tableRow
     * @param posDest
     * @return
     */
    public static XWPFTableCell addTableCell( XWPFTableRow tableRow, int posDest )
    {
        CTTc cTTc = tableRow.getCtRow( ).insertNewTc( posDest );
        XWPFTableCell tableCell = new XWPFTableCell( cTTc, tableRow, tableRow.getTable( ).getBody( ) );
        tableRow.getTableCells( ).add( posDest, tableCell );
        return tableRow.getCell( posDest );
    }

    /**
     * Remove a table cell
     * 
     * @param tableRow
     * @param posDest
     * @return
     */
    public static boolean removeTableCell( XWPFTableRow tableRow, int posDest )
    {
        if ( posDest >= 0 && posDest < tableRow.getTableCells( ).size( ) )
        {
            tableRow.getCtRow( ).removeTc( posDest );
            tableRow.getTableCells( ).remove( posDest );
            return true;
        }
        return false;
    }

    /**
     * Remove an IBodyElement
     * 
     * @param body
     * @param posDest
     */
    public static void removeBodyElement( IBody body, int posDest )
    {
        IBodyElement bodyElement = body.getBodyElements( ).get( posDest );
        XmlCursor cursor = getCursor( bodyElement );
        removeElement( cursor );
    }

    /**
     * Remove an element
     * 
     * @param cursor
     */
    public static void removeElement( XmlCursor cursor )
    {
        cursor.removeXml( );
    }

    /**
     * Get a cursor at position the position of the IBodyElement
     * 
     * @param bodyElement
     * @return
     */
    public static XmlCursor getCursor( IBodyElement bodyElement )
    {
        XmlCursor cursor = null;

        switch( bodyElement.getElementType( ) )
        {
            case PARAGRAPH:
                cursor = ( (XWPFParagraph) bodyElement ).getCTP( ).newCursor( );
                break;
            case TABLE:
                cursor = ( (XWPFTable) bodyElement ).getCTTbl( ).newCursor( );
                break;
        }

        return cursor;
    }
}
