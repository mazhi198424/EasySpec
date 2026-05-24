import pytest
import os


@pytest.fixture
def visiondemo_path():
    """VisionDemo プロジェクトの絶対パス"""
    return os.path.join(os.path.dirname(__file__), '..', '..', 'VisionDemo')


@pytest.fixture
def sample_pom_content():
    return '''<?xml version="1.0" encoding="UTF-8"?>
<project xmlns="http://maven.apache.org/POM/4.0.0">
    <modelVersion>4.0.0</modelVersion>
    <parent>
        <groupId>org.springframework.boot</groupId>
        <artifactId>spring-boot-starter-parent</artifactId>
        <version>3.4.5</version>
    </parent>
    <groupId>com.visiondemo</groupId>
    <artifactId>vision-demo</artifactId>
    <version>1.0.0</version>
    <packaging>war</packaging>
    <name>VisionDemo</name>
</project>'''


@pytest.fixture
def sample_schema_content():
    return '''CREATE TABLE IF NOT EXISTS departments (
    id          BIGINT AUTO_INCREMENT PRIMARY KEY,
    code        VARCHAR(20)  NOT NULL UNIQUE,
    name        VARCHAR(100) NOT NULL,
    parent_id   BIGINT NULL,
    sort_order  INT DEFAULT 0,
    created_at  TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (parent_id) REFERENCES departments(id)
);

CREATE TABLE IF NOT EXISTS employees (
    id              BIGINT AUTO_INCREMENT PRIMARY KEY,
    employee_no     VARCHAR(20)  NOT NULL UNIQUE,
    name            VARCHAR(100) NOT NULL,
    department_id   BIGINT,
    FOREIGN KEY (department_id) REFERENCES departments(id)
);'''


@pytest.fixture
def sample_entity_content():
    return '''@Entity
@Table(name = "employees")
public class Employee {
    @Id
    @GeneratedValue(strategy = GenerationType.IDENTITY)
    private Long id;

    @Column(name = "employee_no", nullable = false, unique = true, length = 20)
    private String employeeNo;

    @Column(nullable = false, length = 100)
    private String name;

    @ManyToOne(fetch = FetchType.LAZY)
    @JoinColumn(name = "department_id")
    private Department department;
}'''


@pytest.fixture
def sample_controller_content():
    return '''@Controller
@RequestMapping("/employee")
public class EmployeeController {

    @Autowired
    private EmployeeService employeeService;

    @GetMapping("/page")
    public String page() {
        return "fragments/employee-list";
    }

    @GetMapping("/api/search")
    @ResponseBody
    public Map<String, Object> search(
            @RequestParam(required = false) String name,
            @RequestParam(required = false) Long deptId,
            @RequestParam(defaultValue = "0") int page,
            @RequestParam(defaultValue = "10") int size) {
        Page<Employee> result = employeeService.search(name, deptId, page, size);
        Map<String, Object> map = new HashMap<>();
        map.put("content", result.getContent());
        map.put("totalPages", result.getTotalPages());
        return map;
    }

    @PostMapping("/api/save")
    @ResponseBody
    public ResponseEntity<Employee> save(@RequestBody Employee employee) {
        Employee saved = employeeService.save(employee);
        return ResponseEntity.ok(saved);
    }

    @DeleteMapping("/api/{id}")
    @ResponseBody
    public ResponseEntity<Void> delete(@PathVariable Long id) {
        employeeService.delete(id);
        return ResponseEntity.ok().build();
    }
}'''


@pytest.fixture
def sample_jsp_content():
    return '''<%@ page contentType="text/html; charset=UTF-8" pageEncoding="UTF-8" %>
<div class="search-bar">
    <label>氏名: <input type="text" id="emp-search-name" size="16"></label>
    <label>部署:
        <select id="emp-search-dept">
            <option value="">すべて</option>
        </select>
    </label>
    <button class="btn btn-primary btn-sm" id="emp-search-btn">検索</button>
    <button class="btn btn-sm" id="emp-reset-btn">リセット</button>
</div>
<div id="emp-table-container"></div>
<div id="emp-pagination-container"></div>
<script>
$(function() {
    function doSearch(page) {
        getJSON('/employee/api/search', {
            name: $('#emp-search-name').val(),
            deptId: $('#emp-search-dept').val(),
            page: page || 0
        }, function(res) {
            TableUtils.renderTable($('#emp-table-container'), {
                columns: [
                    { key: 'employeeNo', label: '社員番号', width: '90px' },
                    { key: 'name', label: '氏名', width: '120px' },
                    { key: 'position', label: '役職', width: '80px' }
                ],
                data: res.content
            });
        });
    }
});
</script>'''
