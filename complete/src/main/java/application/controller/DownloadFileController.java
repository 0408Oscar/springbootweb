package application.controller;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;

@Controller
public class DownloadFileController {
	
	@RequestMapping(value = "/downloadpage")
    public String download(@RequestParam(name="downloadpage", required=false, defaultValue="default downloadpage") String name, Model model) {
        model.addAttribute("downloadpage", name);
        return "/content/downloadpage";
    }
}
