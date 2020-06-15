package com.fine.overtime.controller;

import com.fine.overtime.domain.OverTimeGroup;
import com.fine.overtime.domain.OverTimePeople;
import com.fine.overtime.domain.OverTimeReceipt;
import com.fine.overtime.repo.GroupRepo;
import com.fine.overtime.repo.PeopleRepo;
import com.fine.overtime.repo.ReceiptRepo;
import com.fine.overtime.util.ExcelOverTime;
import com.fine.overtime.util.ExcelRead;
import com.fine.overtime.util.ExcelUploadUtil;
import lombok.RequiredArgsConstructor;
import org.springframework.data.domain.Sort;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import org.springframework.web.servlet.mvc.support.RedirectAttributes;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.File;
import java.util.Comparator;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

@Controller
@RequiredArgsConstructor
public class OverTimeController {

    private final ExcelOverTime overTime;
    private final GroupRepo repo;
    private final ReceiptRepo receiptRepo;
    private final PeopleRepo peopleRepo;

    @RequestMapping(value = "excelUploadAjax", method = RequestMethod.POST)
    public @ResponseBody List<OverTimeReceipt> excelUploadAjax(MultipartHttpServletRequest request, HttpSession httpSession, Model model) throws Exception {

        ExcelUploadUtil util = new ExcelUploadUtil();
        int result = 0;

        MultipartFile excelFile = request.getFile("excelFile");
        Long groupId = Long.valueOf(request.getParameter("groupId"));

        OverTimeGroup group = repo.findById(groupId).get();

        if (excelFile == null || excelFile.isEmpty()) {
            return null;
        } else {
            try {
                model.addAttribute("excelFile", excelFile);
                model.addAttribute("httpSession", httpSession);
                File destFile = util.execute(model);

                ExcelRead c = new ExcelRead();
                List<OverTimeReceipt> list = c.excelUpload(destFile, group);

                receiptRepo.deleteByGroup(groupId);
                receiptRepo.saveAll(list);

                destFile.delete();

                return list;

            } catch (Exception e) {
                // TODO: handle exception
                e.printStackTrace();

                return null;
            }
        }
    }

    @PostMapping("/save")
    public void save(Long groupId, HttpSession session, HttpServletRequest request, HttpServletResponse response) throws Exception {
        OverTimeGroup group = repo.findById(groupId).get();
        overTime.execute(group, session, request, response);
    }

    @GetMapping("/uploadExcel")
    public void uploadExcel(Model model) {

    }

    @GetMapping("/")
    public String index() {

        return "/index";
    }

    @GetMapping("/list")
    public void list(Model model) {

        Sort sort = Sort.by(Sort.Direction.DESC, "groupId");
//        List<OverTimeGroup> list = repo.findAll(sort);
        List<OverTimeGroup> list = repo.findAll();

        model.addAttribute("groupList", list);
    }

    @GetMapping("/createGroup")
    public void createGroup() {

    }

    @PostMapping("/createGroup")
    public String createGroupSubmit(RedirectAttributes rttr, OverTimeGroup group) {

        repo.save(group);

        return redirectAttr(rttr,"success","과제 생성이 완료되었습니다.", "/list");
    }

    @PostMapping("/modifyGroup")
    public void modifyGroup(Long groupId, Model model) {
        getGroupSortList(groupId, model);
    }

    @PostMapping("/modifyGroupSubmit")
    public String modifyGroupSubmit(RedirectAttributes rttr, OverTimeGroup group) {

        Long groupId = group.getGroupId();

        OverTimeGroup beforeGroup = repo.findById(groupId).get();
        group.setReceiptList(beforeGroup.getReceiptList());

        repo.save(group);
        peopleRepo.deleteByGroupIsNull();

        return redirectAttr(rttr,"success","과제 수정이 완료되었습니다.", "/list");
    }

    @PostMapping("/deleteGroup")
    public String deleteGroup(RedirectAttributes rttr, OverTimeGroup group) {

        Long groupId = group.getGroupId();

        repo.deleteById(groupId);

        return redirectAttr(rttr,"success","과제 삭제가 완료되었습니다.", "/list");
    }

    @GetMapping("/detail")
    public void detail(Long groupId, Model model) {
        getGroupSortList(groupId, model);
    }

    private void getGroupSortList(Long groupId, Model model) {
        repo.findById(groupId).ifPresent(dto -> {
            dto.setReceiptList(dto.getReceiptList().stream().sorted(Comparator.comparing(OverTimeReceipt::getReceipt_id)).collect(Collectors.toList()));
            dto.setPeopleList(dto.getPeopleList().stream().sorted(Comparator.comparing(OverTimePeople::getPeople_id)).collect(Collectors.toList()));
            model.addAttribute("dto", dto);
        });
    }

    @GetMapping("/redirectAlert")
    public void redirectAlert() {}

    public String redirectAttr(RedirectAttributes rttr, String status, String msg, String url) {
        Map<String, Object> map = new HashMap<String, Object>();
        map.put("status", status);
        map.put("mmm", msg);
        map.put("url", url);
        rttr.addFlashAttribute("rttrMap", map);
        return "redirect:/redirectAlert";
    }
}
