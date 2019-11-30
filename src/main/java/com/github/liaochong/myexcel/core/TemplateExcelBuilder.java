/*
 * Copyright 2019 liaochong
 *
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.github.liaochong.myexcel.core;

import lombok.extern.slf4j.Slf4j;

import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;

/**
 * 模板Excel构建综合
 *
 * @author liaochong
 * @version 1.0
 */
@Slf4j
public class TemplateExcelBuilder {

    private static String directory;

    private static Map<String, Class<? extends ExcelBuilder>> excelBuilderMapping = new HashMap<>();

    static {
        excelBuilderMapping.put(".ftl", FreemarkerExcelBuilder.class);
        excelBuilderMapping.put(".ttl", ThymeleafExcelBuilder.class);
        excelBuilderMapping.put(".btl", BeetlExcelBuilder.class);
        excelBuilderMapping.put(".tpl", GroovyExcelBuilder.class);
    }

    /**
     * 设置模板所在目录，当前仅FreemarkerExcelBuilder支持
     *
     * @param directory 目录路径
     */
    public synchronized static void directory(String directory) {
        TemplateExcelBuilder.directory = directory;
    }

    /**
     * 设置模板后缀与ExcelBuilder映射
     *
     * @param excelBuilderMapping 映射
     */
    public synchronized static void excelBuilderMapping(Map<String, Class<? extends ExcelBuilder>> excelBuilderMapping) {
        if (excelBuilderMapping == null || excelBuilderMapping.isEmpty()) {
            throw new IllegalArgumentException("ExcelBuilderMap can not be empty");
        }
        TemplateExcelBuilder.excelBuilderMapping = excelBuilderMapping;
    }

    /**
     * 添加一个映射
     *
     * @param suffix     模板后缀
     * @param buildClass ExcelBuilder
     */
    public synchronized static void addExcelBuilderMapping(String suffix, Class<? extends ExcelBuilder> buildClass) {
        excelBuilderMapping.put(suffix, buildClass);
    }

    public static ExcelBuilder template(String path) {
        int index = path.lastIndexOf(".");
        String suffix = path.substring(index);
        Class<? extends ExcelBuilder> builderClass = excelBuilderMapping.get(suffix);
        if (builderClass == null) {
            String suffixList = excelBuilderMapping.keySet().stream().collect(Collectors.joining(","));
            throw new IllegalArgumentException("Please check the template file suffix. The current suffix is not in the list:" + suffixList);
        }
        ExcelBuilder excelBuilder;
        try {
            excelBuilder = builderClass.newInstance();
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        if (directory != null) {
            excelBuilder.directory(directory);
        }
        excelBuilder.template(path);
        return excelBuilder;
    }
}
