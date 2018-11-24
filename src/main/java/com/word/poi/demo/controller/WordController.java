package com.word.poi.demo.controller;

import com.word.poi.demo.service.WordService;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

@RestController
@RequestMapping("/word")
@CrossOrigin(origins = "*", maxAge = 3600)
@Api(tags = "word文档处理")
public class WordController {
    @Autowired
    private WordService wordService;

    @RequestMapping(value = "/exportWord",method = RequestMethod.GET)
    @ApiOperation(value = "生成并导出word文档", httpMethod = "GET", notes = "生成并导出word文档", produces = MediaType.APPLICATION_JSON_UTF8_VALUE)
    public void exportFDDBrief(
            HttpServletResponse response
            ) throws Exception {
        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd-HHmmss");
        //获取统计时间段
        String time = "10.09-10.19";
        String wordName = "Word版-("+time+")"+sdf.format(new Date());
        // 告诉浏览器用什么软件可以打开此文件
        response.reset();// 清空输出流
        response.setHeader("content-Type", "application/msword");
        // 下载文件的默认名称
        response.setHeader("Content-Disposition", "attachment;filename="
                +new String(wordName.getBytes("GB2312"), "8859_1") + ".docx");
        OutputStream out = response.getOutputStream();
        //读取word模板
        InputStream in = this.getClass().getResourceAsStream("/templates/wordTemplate.docx");
        wordService.exportWord(in,out,"全球");

    }

}
