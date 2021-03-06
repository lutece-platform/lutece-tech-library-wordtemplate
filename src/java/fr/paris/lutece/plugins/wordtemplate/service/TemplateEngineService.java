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
import freemarker.core.Environment;
import freemarker.template.Configuration;
import freemarker.template.Template;
import freemarker.template.TemplateException;
import java.io.IOException;
import java.io.StringWriter;

/**
 * Template service based on the Freemarker template engine
 */
public class TemplateEngineService
{

    /**
     * Process the template transformation and return the {@link HtmlTemplate}
     *
     * @param strTemplate
     *            The template name to call
     * @return The {@link HtmlTemplate}
     */
    public static Template createTemplate( String strTemplate )
    {
        Template ftl;
        try
        {
            Configuration cfg = new Configuration( Configuration.VERSION_2_3_28 );
            ftl = new Template( "WordElementTemplate", strTemplate, cfg );
        }
        catch( IOException e )
        {
            throw new WordTemplateException( e.getMessage( ), e );
        }
        return ftl;
    }

    /**
     * Process the template transformation and return the {@link HtmlTemplate}
     *
     * @param strTemplate
     *            The template name to call
     * @param rootMap
     *            The HashMap model
     * @return The {@link HtmlTemplate}
     */
    public static String processTemplate( String strTemplate, Object rootMap )
    {
        StringWriter writer = new StringWriter( 1024 );

        try
        {
            Template template = createTemplate( strTemplate );
            template.process( rootMap, writer );
        }
        catch( IOException | TemplateException e )
        {
            throw new WordTemplateException( e.getMessage( ), e );
        }
        return writer.toString( );
    }

    /**
     * Process the template transformation and return the {@link HtmlTemplate}
     *
     * @param strTemplate
     *            The template name to call
     * @param rootMap
     *            The HashMap model
     * @return The {@link HtmlTemplate}
     */
    public static Environment getEnvironment( String strTemplate, Object rootMap )
    {
        StringWriter writer = new StringWriter( 1024 );

        Environment environment;
        try
        {
            Template template = createTemplate( strTemplate );
            environment = template.createProcessingEnvironment( rootMap, writer );
            environment.process( );
        }
        catch( IOException | TemplateException e )
        {
            throw new WordTemplateException( e.getMessage( ), e );
        }
        return environment;
    }

}
