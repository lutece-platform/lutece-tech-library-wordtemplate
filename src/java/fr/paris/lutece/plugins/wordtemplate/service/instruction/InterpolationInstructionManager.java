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
package fr.paris.lutece.plugins.wordtemplate.service.instruction;

import fr.paris.lutece.plugins.wordtemplate.business.IWordTemplateElement;
import fr.paris.lutece.plugins.wordtemplate.business.InterpolationInstructionElement;
import fr.paris.lutece.plugins.wordtemplate.service.TemplateEngineService;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Service that manage interpolation instructions
 */
public class InterpolationInstructionManager implements IInstructionManager
{
    private static final String INTERPOLATION_PATTERN = "\\$\\{.*?\\}";

    /**
     *
     * @return
     */
    public static boolean isMulti( )
    {
        return false;
    }

    /**
     *
     * @param strExpression
     * @return
     */
    public boolean isOfType( String strExpression )
    {
        Pattern pattern = Pattern.compile( INTERPOLATION_PATTERN );
        Matcher matcher = pattern.matcher( strExpression );
        return matcher.find( );
    }

    /**
     *
     * @param strExpression
     * @param run
     * @return
     */
    @Override
    public IWordTemplateElement createInstruction( String strExpression, XWPFRun run )
    {
        return new InterpolationInstructionElement( strExpression, run );
    }

    /**
     *
     * @param element
     * @param model
     */
    @Override
    public void processInstruction( IWordTemplateElement element, Map<String, Object> model )
    {
        if ( element instanceof InterpolationInstructionElement )
        {
            InterpolationInstructionElement instruction = (InterpolationInstructionElement) element;
            XWPFRun run = instruction.getRun( );
            run.setText( evaluateExpression( instruction.getExpression( ), model ), 0 );
        }
    }

    /**
     *
     * @param strExpression
     * @param model
     * @return
     */
    private String evaluateExpression( String strExpression, Map<String, Object> model )
    {
        return TemplateEngineService.processTemplate( strExpression, model );
    }
}
