package com.fr.plugin.http;

import com.fr.decision.fun.HttpHandler;
import com.fr.decision.fun.impl.AbstractHttpHandlerProvider;

/**
 * 提供http服务
 */
public class HttpClient extends AbstractHttpHandlerProvider {
    @Override
    public HttpHandler[] registerHandlers() {
        return new HttpHandler[]{
             new HttpExportExcel()
        };
    }
}
