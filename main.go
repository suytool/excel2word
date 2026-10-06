package main

import (
	"archive/zip"
	"bytes"
	"encoding/json"
	"encoding/xml"
	"fmt"
	"io"
	"log"
	"math/rand/v2"
	"net/http"
	"net/url"
	"os"
	"path/filepath"
	"regexp"
	"strconv"
	"strings"
	"time"

	"github.com/xuri/excelize/v2"
)

// ============ Excel 解析 ============

type ExcelData struct {
	Headers []string
	Rows    [][]string
}

func readExcel(path string, headerRow int, startRow int, endRow int) (*ExcelData, error) {
	f, err := excelize.OpenFile(path)
	if err != nil {
		return nil, fmt.Errorf("打开 Excel 失败: %v", err)
	}
	defer f.Close()

	sheetName := f.GetSheetName(0)

	// 先获取表头行，确定列数
	hr := headerRow
	if hr < 1 {
		hr = 1
	}

	// 用 rows 迭代器获取总行数和表头列数
	rows, err := f.GetRows(sheetName)
	if err != nil {
		return nil, fmt.Errorf("读取 Excel 失败: %v", err)
	}
	if len(rows) == 0 {
		return nil, fmt.Errorf("Excel 为空")
	}

	hrIdx := hr - 1
	if hrIdx >= len(rows) {
		return nil, fmt.Errorf("表头行超出范围")
	}

	// 表头
	headers := make([]string, 0, len(rows[hrIdx]))
	for _, h := range rows[hrIdx] {
		headers = append(headers, strings.TrimSpace(h))
	}
	numCols := len(headers)

	// 数据范围（1-based）
	sr := startRow
	if sr <= hr {
		sr = hr + 1
	}
	er := endRow
	if er > len(rows) || endRow == 0 {
		er = len(rows)
	}

	var data [][]string
	for r := sr; r <= er && r <= len(rows); r++ {
		// 用 GetCellValue 逐列读取，确保列对齐
		cells := make([]string, numCols)
		for c := 0; c < numCols; c++ {
			colName, _ := excelize.ColumnNumberToName(c + 1)
			cellRef := fmt.Sprintf("%s%d", colName, r)
			val, err := f.GetCellValue(sheetName, cellRef)
			if err == nil {
				cells[c] = val
			}
		}
		// 跳过完全空行
		isEmpty := true
		for _, c := range cells {
			if strings.TrimSpace(c) != "" {
				isEmpty = false
				break
			}
		}
		if isEmpty {
			continue
		}
		data = append(data, cells)
	}

	if len(data) == 0 {
		return nil, fmt.Errorf("指定范围内没有数据行")
	}

	return &ExcelData{Headers: headers, Rows: data}, nil
}

// ============ Word 模板替换 ============

var placeholderRegex = regexp.MustCompile(`\{\{\s*([^}]+?)\s*\}\}`)

func processDocx(tplPath string, data map[string]string) ([]byte, error) {
	r, err := zip.OpenReader(tplPath)
	if err != nil {
		return nil, fmt.Errorf("打开 Word 模板失败: %v", err)
	}
	defer r.Close()

	buf := new(bytes.Buffer)
	w := zip.NewWriter(buf)

	for _, file := range r.File {
		rc, err := file.Open()
		if err != nil {
			return nil, err
		}
		content, err := io.ReadAll(rc)
		rc.Close()
		if err != nil {
			return nil, err
		}

		if file.Name == "word/document.xml" {
			content = replacePlaceholders(content, data)
		}

		fw, err := w.Create(file.Name)
		if err != nil {
			return nil, err
		}
		if _, err := fw.Write(content); err != nil {
			return nil, err
		}
	}
	w.Close()
	return buf.Bytes(), nil
}

func replacePlaceholders(xmlData []byte, data map[string]string) []byte {
	xmlStr := string(xmlData)
	xmlStr = mergeSplitPlaceholders(xmlStr)
	xmlStr = placeholderRegex.ReplaceAllStringFunc(xmlStr, func(match string) string {
		sub := placeholderRegex.FindStringSubmatch(match)
		if len(sub) < 2 {
			return match
		}
		key := strings.TrimSpace(sub[1])
		if val, ok := data[key]; ok {
			return xmlEscape(val)
		}
		return ""
	})
	return []byte(xmlStr)
}

// mergeSplitPlaceholders 只合并那些占位符确实被 Word 分割到多个 run 的情况，
// 不会合并普通文本 run，从而保留原有字体和格式。
func mergeSplitPlaceholders(xmlStr string) string {
	type textNode struct {
		cStart  int // 文本内容开始（<w:t...> 的 > 之后）
		cEnd    int // 文本内容结束（</w:t> 之前）
		content string
	}
	textRe := regexp.MustCompile(`<w:t[^>]*>.*?</w:t>`)
	var nodes []textNode
	for _, m := range textRe.FindAllStringSubmatchIndex(xmlStr, -1) {
		fs, fe := m[0], m[1]
		tag := xmlStr[fs:fe]
		gt := strings.Index(tag, ">")
		cs := fs + gt + 1
		ce := fe - len("</w:t>")
		nodes = append(nodes, textNode{cs, ce, xmlStr[cs:ce]})
	}
	if len(nodes) < 2 {
		return xmlStr
	}

	type del struct{ start, end int }
	var dels []del
	depth := 0
	for i := 0; i < len(nodes); i++ {
		c := nodes[i].content
		if depth > 0 && i > 0 {
			// 合并：保留前一个 <w:t>开标签和部分1文本，
			// 删除 </w:t></w:r><w:r><w:rPr>...</w:rPr><w:t>，
			// 保留部分2文本和 </w:t></w:r>
			dels = append(dels, del{nodes[i-1].cEnd, nodes[i].cStart})
		}
		depth = scanPlaceholderDepth(c, depth)
	}

	for i := len(dels) - 1; i >= 0; i-- {
		d := dels[i]
		xmlStr = xmlStr[:d.start] + xmlStr[d.end:]
	}
	return xmlStr
}

// scanPlaceholderDepth 从当前 depth 开始扫描 s，返回扫描后的 depth。
func scanPlaceholderDepth(s string, start int) int {
	depth := start
	i := 0
	for i < len(s)-1 {
		if s[i] == '{' && s[i+1] == '{' {
			depth++
			i += 2
		} else if s[i] == '}' && s[i+1] == '}' {
			depth--
			i += 2
		} else {
			i++
		}
	}
	return depth
}

func xmlEscape(s string) string {
	var b bytes.Buffer
	xml.EscapeText(&b, []byte(s))
	return b.String()
}

// sanitizeFilename 清理文件名中的非法字符
func sanitizeFilename(s string) string {
	s = strings.Map(func(r rune) rune {
		if strings.ContainsRune(`<>:"/\|?*`, r) {
			return '_'
		}
		return r
	}, s)
	if len(s) > 80 {
		s = s[:80]
	}
	return s
}

// renderFilename 用数据渲染文件名模板 {{列名}}
func renderFilename(tpl string, row map[string]string) string {
	result := placeholderRegex.ReplaceAllStringFunc(tpl, func(match string) string {
		sub := placeholderRegex.FindStringSubmatch(match)
		if len(sub) < 2 {
			return match
		}
		key := strings.TrimSpace(sub[1])
		if val, ok := row[key]; ok {
			return sanitizeFilename(val)
		}
		return ""
	})
	return result
}

// ============ HTTP 服务 ============

// staticDir 定位前端静态目录：优先 STATIC_DIR 环境变量，
// 其次根据可执行文件位置推断（兼容 bin/ 与根目录两种布局），不依赖进程工作目录。
func staticDir() string {
	if d := os.Getenv("STATIC_DIR"); d != "" {
		return d
	}
	if exe, err := os.Executable(); err == nil {
		dir := filepath.Dir(exe)
		for _, p := range []string{filepath.Join(dir, "static"), filepath.Join(dir, "..", "static")} {
			if fi, err := os.Stat(p); err == nil && fi.IsDir() {
				return p
			}
		}
	}
	return "static"
}

// outputDir 生成的 Word 输出目录，可用 OUTPUT_DIR 覆盖（fnOS 下指向应用数据目录以持久化）。
func outputDir() string {
	if d := os.Getenv("OUTPUT_DIR"); d != "" {
		return d
	}
	return "/tmp/excel2word_output"
}

// uploadDir 上传的临时文件目录，可用 UPLOAD_DIR 覆盖。
func uploadDir() string {
	if d := os.Getenv("UPLOAD_DIR"); d != "" {
		return d
	}
	return "/tmp"
}

// tempFilePath 生成带时间戳与随机后缀的唯一临时文件路径，避免并发请求互相覆盖。
func tempFilePath(prefix, suffix string) string {
	dir := uploadDir()
	os.MkdirAll(dir, 0755)
	return filepath.Join(dir, fmt.Sprintf("%s_%d_%d%s", prefix, time.Now().UnixNano(), rand.IntN(1<<30), suffix))
}

// contentDisposition 生成下载响应头，中文文件名使用 RFC 5987 filename* 编码。
func contentDisposition(name string) string {
	ascii := true
	for _, r := range name {
		if r > 127 {
			ascii = false
			break
		}
	}
	if ascii {
		return `attachment; filename="` + name + `"`
	}
	return `attachment; filename="download.docx"; filename*=UTF-8''` + url.PathEscape(name)
}

func main() {
	port := os.Getenv("PORT")
	if port == "" {
		port = "7721"
	}

	http.Handle("/", http.FileServer(http.Dir(staticDir())))

	// 解析 Excel 预览
	http.HandleFunc("/api/parse", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}
		excelFile, _, err := r.FormFile("excel")
		if err != nil {
			http.Error(w, "缺少 Excel 文件", 400)
			return
		}
		defer excelFile.Close()
		tmpExcel := tempFilePath("upload_excel", ".xlsx")
		f, _ := os.Create(tmpExcel)
		io.Copy(f, excelFile)
		f.Close()

		headerRow, _ := strconv.Atoi(r.FormValue("headerRow"))
		if headerRow == 0 {
			headerRow = 1
		}

		data, err := readExcel(tmpExcel, headerRow, 1, 0)
		if err != nil {
			http.Error(w, err.Error(), 500)
			return
		}

		// 返回表头和总行数，以及前5行预览
		type PreviewRow map[string]string
		type ParseResp struct {
			Headers     []string     `json:"headers"`
			TotalRows   int          `json:"totalRows"`
			PreviewRows []PreviewRow `json:"previewRows"`
		}
		resp := ParseResp{
			Headers:   data.Headers,
			TotalRows: len(data.Rows),
		}
		n := min(5, len(data.Rows))
		for i := 0; i < n; i++ {
			m := make(PreviewRow)
			for ci, h := range data.Headers {
				if ci < len(data.Rows[i]) {
					m[h] = data.Rows[i][ci]
				} else {
					m[h] = ""
				}
			}
			resp.PreviewRows = append(resp.PreviewRows, m)
		}
		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(resp)
	})

	// 生成 Word
	outDir := outputDir()
	os.MkdirAll(outDir, 0755)

	http.HandleFunc("/api/generate", func(w http.ResponseWriter, r *http.Request) {
		if r.Method != http.MethodPost {
			http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
			return
		}

		excelFile, _, err := r.FormFile("excel")
		if err != nil {
			http.Error(w, "缺少 Excel 文件", 400)
			return
		}
		defer excelFile.Close()
		tmpExcel := tempFilePath("gen_excel", ".xlsx")
		f, _ := os.Create(tmpExcel)
		io.Copy(f, excelFile)
		f.Close()

		wordFile, _, err := r.FormFile("word")
		if err != nil {
			http.Error(w, "缺少 Word 模板", 400)
			return
		}
		defer wordFile.Close()
		tmpWord := tempFilePath("gen_template", ".docx")
		f2, _ := os.Create(tmpWord)
		io.Copy(f2, wordFile)
		f2.Close()

		headerRow, _ := strconv.Atoi(r.FormValue("headerRow"))
		if headerRow == 0 {
			headerRow = 1
		}
		startRow, _ := strconv.Atoi(r.FormValue("startRow"))
		endRow, _ := strconv.Atoi(r.FormValue("endRow"))
		rowsPerDoc, _ := strconv.Atoi(r.FormValue("rowsPerDoc"))
		if rowsPerDoc <= 0 {
			rowsPerDoc = 1
		}
		nameTemplate := r.FormValue("nameTemplate")
		randomOrder := r.FormValue("randomOrder") == "true"

		data, err := readExcel(tmpExcel, headerRow, startRow, endRow)
		if err != nil {
			http.Error(w, err.Error(), 500)
			return
		}

		rows := data.Rows
		headers := data.Headers

		if randomOrder {
			rand.Shuffle(len(rows), func(i, j int) {
				rows[i], rows[j] = rows[j], rows[i]
			})
		}

		// 按 rowsPerDoc 分组
		type Group struct {
			rows []map[string]string
		}
		var groups []Group
		for i := 0; i < len(rows); i += rowsPerDoc {
			end := i + rowsPerDoc
			if end > len(rows) {
				end = len(rows)
			}
			var grp Group
			for j := i; j < end; j++ {
				m := make(map[string]string)
				for ci, h := range headers {
					if ci < len(rows[j]) {
						m[h] = rows[j][ci]
					} else {
						m[h] = ""
					}
				}
				grp.rows = append(grp.rows, m)
			}
			groups = append(groups, grp)
		}

		// 清空输出目录
		os.RemoveAll(outDir)
		os.MkdirAll(outDir, 0755)

		type FileInfo struct {
			Name string `json:"name"`
			Size int64  `json:"size"`
		}
		var files []FileInfo

		for gi, grp := range groups {
			var docBytes []byte
			var fname string
			if len(grp.rows) == 1 {
				docBytes, err = processDocx(tmpWord, grp.rows[0])
				if err != nil {
					log.Printf("第 %d 个文档失败: %v", gi+1, err)
					continue
				}
				fname = genFilename(nameTemplate, grp.rows, headers, gi)
			} else {
				docBytes, err = mergeDocxWithRows(tmpWord, grp.rows)
				if err != nil {
					log.Printf("第 %d 组合并失败: %v", gi+1, err)
					continue
				}
				fname = genFilename(nameTemplate, grp.rows, headers, gi)
			}
			outPath := outDir + "/" + fname
			os.WriteFile(outPath, docBytes, 0644)
			files = append(files, FileInfo{Name: fname, Size: int64(len(docBytes))})
		}

		w.Header().Set("Content-Type", "application/json")
		json.NewEncoder(w).Encode(map[string]interface{}{
			"files": files,
			"count": len(files),
		})
	})

	// 单独下载
	http.HandleFunc("/api/download", func(w http.ResponseWriter, r *http.Request) {
		fname := r.URL.Query().Get("file")
		if fname == "" || strings.Contains(fname, "..") || strings.Contains(fname, "/") {
			http.Error(w, "invalid filename", 400)
			return
		}
		path := outDir + "/" + fname
		w.Header().Set("Content-Type", "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
		w.Header().Set("Content-Disposition", contentDisposition(fname))
		http.ServeFile(w, r, path)
	})

	// 打包下载全部
	http.HandleFunc("/api/download-all", func(w http.ResponseWriter, r *http.Request) {
		entries, _ := os.ReadDir(outDir)
		w.Header().Set("Content-Type", "application/zip")
		w.Header().Set("Content-Disposition", `attachment; filename="generated.zip"`)
		zipW := zip.NewWriter(w)
		for _, e := range entries {
			if e.IsDir() {
				continue
			}
			data, err := os.ReadFile(outDir + "/" + e.Name())
			if err != nil {
				continue
			}
			fw, _ := zipW.Create(e.Name())
			fw.Write(data)
		}
		zipW.Close()
	})

	log.Printf("Excel2Word 服务启动，端口 %s，静态目录 %s，输出目录 %s", port, staticDir(), outDir)
	log.Fatal(http.ListenAndServe(":"+port, nil))
}

func genFilename(tpl string, rows []map[string]string, headers []string, idx int) string {
	first := rows[0]
	var base string

	// 从模板中提取第一个列名
	tplCol := ""
	if tpl != "" {
		base = renderFilename(tpl, first)
		if m := placeholderRegex.FindStringSubmatch(tpl); len(m) >= 2 {
			tplCol = strings.TrimSpace(m[1])
		}
	}
	if base == "" && len(headers) > 0 {
		base = first[headers[0]]
		tplCol = headers[0]
	}
	if base == "" {
		base = "document"
	}
	base = sanitizeFilename(base)

	// 多行合并：文件名用 第一个~最后一个（用同一个列名）
	if len(rows) > 1 && tplCol != "" {
		last := rows[len(rows)-1]
		lastVal := sanitizeFilename(last[tplCol])
		if lastVal != "" && lastVal != base {
			return fmt.Sprintf("%s~%s.docx", base, lastVal)
		}
	}
	return fmt.Sprintf("%s_%d.docx", base, idx+1)
}

// mergeDocxWithRows 将模板内容重复N次，中间加分页符
func mergeDocxWithRows(tplPath string, rows []map[string]string) ([]byte, error) {
	// 读取模板
	r, err := zip.OpenReader(tplPath)
	if err != nil {
		return nil, err
	}
	defer r.Close()

	// 读取 document.xml
	var docXml string
	var otherFiles = map[string][]byte{}
	for _, file := range r.File {
		content, _ := func() ([]byte, error) {
			rc, err := file.Open()
			if err != nil {
				return nil, err
			}
			defer rc.Close()
			return io.ReadAll(rc)
		}()
		if file.Name == "word/document.xml" {
			docXml = string(content)
		} else {
			otherFiles[file.Name] = content
		}
	}

	// 提取 body 内容（不含 sectPr）
	bodyStart := strings.Index(docXml, "<w:body>") + len("<w:body>")
	sectStart := strings.Index(docXml, "<w:sectPr")
	if sectStart < 0 {
		sectStart = strings.Index(docXml, "</w:body>")
	}
	bodyContent := docXml[bodyStart:sectStart]
	sectPr := docXml[sectStart:]
	// sectPr 到 </w:body>
	endBody := strings.Index(sectPr, "</w:body>")
	if endBody >= 0 {
		sectPr = sectPr[:endBody]
	}

	// 分页符 XML
	pageBreak := `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`

	// 为每行替换模板
	var combinedBody string
	for i, row := range rows {
		replaced := replacePlaceholders([]byte(bodyContent), row)
		combinedBody += string(replaced)
		if i < len(rows)-1 {
			combinedBody += pageBreak
		}
	}

	newDocXml := docXml[:bodyStart] + combinedBody + sectPr + "</w:body>" + docXml[strings.Index(docXml, "</w:body>")+len("</w:body>"):]

	buf := new(bytes.Buffer)
	w := zip.NewWriter(buf)
	for name, content := range otherFiles {
		fw, _ := w.Create(name)
		fw.Write(content)
	}
	fw, _ := w.Create("word/document.xml")
	fw.Write([]byte(newDocXml))
	w.Close()
	return buf.Bytes(), nil
}
