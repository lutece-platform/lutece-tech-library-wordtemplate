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

import fr.paris.lutece.plugins.wordtemplate.business.IWordTemplateElement;
import fr.paris.lutece.plugins.wordtemplate.service.instruction.IInstructionManager;
import fr.paris.lutece.plugins.wordtemplate.service.instruction.InterpolationInstructionManager;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 * Service for instruction management
 */
public class InstructionService
{
    private static InstructionService _instructionService;
    private List<IInstructionManager> _listInstructionManager;

    private InstructionService( )
    {
        _listInstructionManager = new ArrayList<>( );
        _listInstructionManager.add( new InterpolationInstructionManager( ) );
    }

    /**
     *
     * @return
     */
    public static InstructionService init( )
    {
        if ( _instructionService == null )
        {
            _instructionService = new InstructionService( );
        }

        return _instructionService;
    }

    /**
     *
     * @param expression
     * @param run
     * @return
     */
    public IWordTemplateElement createInstruction( String expression, XWPFRun run )
    {
        for ( IInstructionManager instructionManager : _listInstructionManager )
        {
            if ( instructionManager.isOfType( expression ) )
            {
                IWordTemplateElement element = instructionManager.createInstruction( expression, run );
                return element;
            }
        }
        return null;
    }

    /**
     *
     * @param instruction
     * @param model
     */
    public void processInstruction( IWordTemplateElement instruction, Map<String, Object> model )
    {
        for ( IInstructionManager instructionManager : _listInstructionManager )
        {
            instructionManager.processInstruction( instruction, model );
        }
    }
}
