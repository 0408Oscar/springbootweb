package application.controller;

import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;

@Controller
public class SayHiController {
	
	@GetMapping("/sayhi")
    public String SayHi(@RequestParam(name="action", required=false, defaultValue="default Say hi") String name, Model model) {
        model.addAttribute("action", name);
        return "/content/sayHi";
    }
}
