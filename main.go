package main

import (
	"context"
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"path/filepath"
	"strings"
	"time"

	"github.com/antchfx/htmlquery"
	"github.com/ncruces/zenity"
	"github.com/xuri/excelize/v2"
	"golang.org/x/mod/modfile"
)

type PackageInfo struct {
	Name            string
	Version         string
	License         string
	LicenseURL      string
	Author          string
	Description     string
	Copyright       string
	PackageURL      string
	GitHubURL       string
	RepositoryType  string
	Repository      string
	ModuleNameNoVer string
}

// Package represents a dependency
type Package struct {
	Path    string
	Version string
	GoMod   bool
}

// Parse go.mod file
func parseGoMod(filename string) ([]Package, string, error) {
	data, err := os.ReadFile(filename)
	if err != nil {
		return nil, "", err
	}

	// Use ParseLax to allow unknown block types
	file, err := modfile.ParseLax(filepath.Base(filename), data, nil)
	if err != nil {
		return nil, "", err
	}

	var packages []Package
	for _, req := range file.Require {
		packages = append(packages, Package{
			Path:    req.Mod.Path,
			Version: req.Mod.Version,
			GoMod:   true,
		})
	}

	// Get module name from the parsed file
	moduleName := file.Module.Mod.Path + "-api"
	return packages, moduleName, nil
}

// Parse package.json file
func parsePackageJSON(filename string) ([]Package, string, error) {
	data, err := os.ReadFile(filename)
	if err != nil {
		return nil, "", err
	}

	var packageJSON struct {
		Name            string            `json:"name"`
		Dependencies    map[string]string `json:"dependencies"`
		DevDependencies map[string]string `json:"devDependencies"`
	}

	if err := json.Unmarshal(data, &packageJSON); err != nil {
		return nil, "", err
	}

	var packages []Package

	for name, version := range packageJSON.Dependencies {
		packages = append(packages, Package{
			Path:    name,
			Version: version,
			GoMod:   false,
		})
	}

	for name, version := range packageJSON.DevDependencies {
		packages = append(packages, Package{
			Path:    name,
			Version: version,
			GoMod:   false,
		})
	}

	return packages, packageJSON.Name + "-ui", nil
}

// Get metadata from pkg.go.dev
func getGoModMetadata(pkg *Package) PackageInfo {
	info := PackageInfo{
		Name:           pkg.Path,
		Version:        pkg.Version,
		PackageURL:     pkg.Path + "/@v/" + pkg.Version + ".info",
		RepositoryType: "go",
	}

	// Create HTTP client with timeout
	client := &http.Client{
		Timeout: 10 * time.Second,
		Transport: &http.Transport{
			MaxIdleConns:          10,
			IdleConnTimeout:       30 * time.Second,
			DisableCompression:    false,
			DisableKeepAlives:     false,
			ResponseHeaderTimeout: 5 * time.Second,
		},
	}

	// Get license and other info from pkg.go.dev
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()

	req, err := http.NewRequestWithContext(ctx, "GET", "https://pkg.go.dev/"+pkg.Path, nil)
	if err != nil {
		return info
	}

	resp, err := client.Do(req)
	if err != nil || resp.StatusCode != 200 {
		return info
	}
	defer resp.Body.Close()

	// Parse HTML from response
	doc, err := htmlquery.Parse(resp.Body)
	if err == nil {
		// Find license
		node := htmlquery.FindOne(doc, `//span[contains(@class, "License")]/a`)
		if node == nil {
			node = htmlquery.FindOne(doc, `//a[contains(@href, "licenses")]`)
		}
		if node == nil {
			node = htmlquery.FindOne(doc, `//span[contains(@class, "license")]`)
		}
		if node != nil {
			txt := strings.TrimSpace(htmlquery.InnerText(node))
			if !strings.Contains(txt, "not legal advice") && txt != "" {
				info.License = txt
				info.LicenseURL = "https://licenses.nuget.org/" + txt
			}
		}

		// Find description
		node = htmlquery.FindOne(doc, `//h2[contains(@class, "package-title")]/following-sibling::p`)
		if node == nil {
			node = htmlquery.FindOne(doc, `//div[contains(@class, "package-details")]/p`)
		}
		if node == nil {
			node = htmlquery.FindOne(doc, `//div[contains(@class, "documentation")]//p`)
		}
		if node == nil {
			node = htmlquery.FindOne(doc, `//div[contains(@class, "pkg-subdoc")]//p`)
		}
		if node != nil {
			info.Description = strings.TrimSpace(htmlquery.InnerText(node))
		}

		// Find repository link (GitHub or other) - try multiple selectors to be more robust
		repositorySelectors := []string{
			`//div[contains(@class, "UnitMeta-repo")]//a`,
			`//html/body/aside/nav/ul/li[5]/div/div/ul/li[3]/a`,
			`//aside//a[contains(@href, ".")]`,
			`//div[contains(@class, "repository")]//a`,
		}

		for _, selector := range repositorySelectors {
			node = htmlquery.FindOne(doc, selector)
			if node != nil {
				url := htmlquery.SelectAttr(node, "href")
				if url != "" && !strings.Contains(url, "pkg.go.dev") {
					info.GitHubURL = url
					break
				}
			}
		}

		// If still no GitHub URL found, try to construct from module path
		if info.GitHubURL == "" && strings.Contains(pkg.Path, "github.com/") {
			info.GitHubURL = "https://" + pkg.Path
		}

		// Try multiple approaches to find author/maintainer info from page
		authorSelectors := []string{
			`//span[contains(@class, "Author")]`,
			`//div[contains(@class, "author")]`,
			`//span[contains(@class, "text-muted")]`,
			`//div[contains(@class, "meta")]//span[not(contains(@class, "license"))]`,
			`//div[contains(@class, "details")]//span[1]`,
			`//div[contains(@class, "pkg-subdoc")]/p/span`,
		}

		for _, selector := range authorSelectors {
			node = htmlquery.FindOne(doc, selector)
			if node != nil {
				author := strings.TrimSpace(htmlquery.InnerText(node))
				if author != "" && !strings.Contains(strings.ToLower(author), "license") &&
					!strings.Contains(strings.ToLower(author), "copyright") && len(author) < 100 {
					info.Author = author
					break
				}
			}
		}

		// If no author found from page, try to infer from package path
		if info.Author == "" {
			// For GitHub repos, extract user/organization name
			if strings.Contains(pkg.Path, "github.com/") {
				parts := strings.Split(pkg.Path, "/")
				if len(parts) >= 2 {
					info.Author = parts[1]
				}
			}
			// Try other common patterns
			if info.Author == "" && strings.Contains(pkg.Path, "/") {
				parts := strings.Split(pkg.Path, "/")
				if len(parts) >= 2 {
					info.Author = parts[0]
				}
			}
		}

		// Try to extract copyright info from license or page
		if info.License != "" {
			info.Copyright = info.License + " Copyright"
		} else {
			// Look for copyright mentions
			node = htmlquery.FindOne(doc, `//span[contains(text(), "Copyright")]`)
			if node == nil {
				node = htmlquery.FindOne(doc, `//div[contains(text(), "©")]`)
			}
			if node == nil {
				node = htmlquery.FindOne(doc, `//span[contains(text(), "©")]`)
			}
			if node != nil {
				copyright := strings.TrimSpace(htmlquery.InnerText(node))
				if copyright != "" {
					info.Copyright = copyright
				}
			}
		}
	}

	return info
}

// Get metadata from npm registry
func getNPMMetadata(pkg *Package) PackageInfo {
	info := PackageInfo{
		Name:            pkg.Path,
		Version:         pkg.Version,
		ModuleNameNoVer: pkg.Path,
		RepositoryType:  "npm",
	}

	// Clean version (remove ^, ~, etc.)
	version := strings.TrimPrefix(strings.TrimPrefix(pkg.Version, "^"), "~")

	// Create HTTP client with timeout
	client := &http.Client{
		Timeout: 10 * time.Second,
		Transport: &http.Transport{
			MaxIdleConns:          10,
			IdleConnTimeout:       30 * time.Second,
			DisableCompression:    false,
			DisableKeepAlives:     false,
			ResponseHeaderTimeout: 5 * time.Second,
		},
	}

	// Get info from npm registry with context
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()

	req, err := http.NewRequestWithContext(ctx, "GET", "https://registry.npmjs.org/"+pkg.Path+"/"+version, nil)
	if err != nil {
		return info
	}

	resp, err := client.Do(req)
	if err == nil && resp.StatusCode == 200 {
		defer resp.Body.Close()
		var npmPkg struct {
			License  string `json:"license"`
			Licenses []struct {
				Type string `json:"type"`
			} `json:"licenses"`
			Author      any                 `json:"author"`
			Maintainers []map[string]string `json:"maintainers"`
			Description string              `json:"description"`
			Repository  struct {
				Type string `json:"type"`
				URL  string `json:"url"`
			} `json:"repository"`
			Homepage string `json:"homepage"`
			Readme   string `json:"readme"`
		}

		if err := json.NewDecoder(resp.Body).Decode(&npmPkg); err == nil {
			// Get license
			if npmPkg.License != "" {
				info.License = npmPkg.License
				info.LicenseURL = "https://licenses.nuget.org/" + npmPkg.License
			} else if len(npmPkg.Licenses) > 0 {
				info.License = npmPkg.Licenses[0].Type
				info.LicenseURL = "https://licenses.nuget.org/" + npmPkg.Licenses[0].Type
			}

			// Get author - try multiple sources
			if author, ok := npmPkg.Author.(map[string]any); ok {
				if name, ok := author["name"]; ok {
					info.Author = name.(string)
				} else if email, ok := author["email"]; ok {
					info.Author = email.(string)
				}
			} else if authorStr, ok := npmPkg.Author.(string); ok && authorStr != "" {
				info.Author = authorStr
			}

			// If no author from main field, try maintainers
			if info.Author == "" && len(npmPkg.Maintainers) > 0 {
				if name, ok := npmPkg.Maintainers[0]["name"]; ok {
					info.Author = name
				} else if email, ok := npmPkg.Maintainers[0]["email"]; ok {
					info.Author = email
				}
			}

			info.Description = npmPkg.Description

			// Get repository/GitHub URL
			if npmPkg.Repository.URL != "" {
				info.Repository = npmPkg.Repository.URL
				info.GitHubURL = npmPkg.Repository.URL
			} else if npmPkg.Homepage != "" {
				info.Repository = npmPkg.Homepage
			}

			// Try to extract copyright from README or license
			if info.License != "" {
				info.Copyright = info.License + " Copyright"
			} else if npmPkg.Readme != "" {
				// Try to find copyright mentions in README
				for line := range strings.SplitSeq(npmPkg.Readme, "\n") {
					if strings.Contains(strings.ToLower(line), "copyright") ||
						strings.Contains(line, "©") {
						info.Copyright = strings.TrimSpace(line)
						break
					}
				}
			}
		}
	}

	return info
}

func main() {
	wd, err := os.Getwd()
	if err != nil {
		zenity.Error("Failed to get current working directory: "+err.Error(), zenity.Title("Error"), zenity.ErrorIcon)
		return
	}

	inName, err := zenity.SelectFile(
		zenity.Filename(wd),
		zenity.FileFilters{
			{
				Name:     "Go Module or Package JSON",
				Patterns: []string{"go.mod", "package.json"},
				CaseFold: false},
		},
	)
	if err != nil {
		// User cancelled - exit process instead of showing error dialog
		os.Exit(1)
	}

	isGoMod := strings.HasSuffix(inName, "go.mod")
	var moduleName string
	var packages []Package

	// Parse file
	if isGoMod {
		packages, moduleName, err = parseGoMod(inName)
	} else {
		packages, moduleName, err = parsePackageJSON(inName)
	}
	if err != nil {
		zenity.Error("Failed to parse file: "+err.Error(), zenity.Title("Error"), zenity.ErrorIcon)
		return
	}

	outName := moduleName + "_license.xlsx"

	dlg, err := zenity.Progress(
		zenity.Title("Running..."))
	if err != nil {
		zenity.Error("Create progress dialog failed: "+err.Error(), zenity.Title("Error"), zenity.ErrorIcon)
		os.Exit(1)
	}
	defer dlg.Close()

	// Create Excel workbook
	f := excelize.NewFile()

	// Get current sheet name
	sheetName := f.GetSheetName(0)

	// Write header based on file type
	header := []string{}
	if isGoMod {
		header = []string{"Name", "License", "PackageVersion", "LicenseURL", "Author", "Description", "Copyright", "PackageURL", "GitHubURL", "RepositoryType"}
	} else {
		header = []string{"Module Name", "License", "Repository", "License URL", "Author", "Description", "Copyright", "GitHub URL", "Module Name (No Version)", "Version"}
	}

	// Write header row
	for i, col := range header {
		cell := fmt.Sprintf("%s1", string(rune('A'+i)))
		f.SetCellValue(sheetName, cell, col)
	}

	total := len(packages)
	for i, pkg := range packages {
		dlg.Value(int(float64(i) / float64(total) * 100))
		dlg.Text("Processing " + pkg.Path + "...")

		var info PackageInfo
		if isGoMod {
			info = getGoModMetadata(&pkg)
			row := []interface{}{
				info.Name,
				info.License,
				info.Version,
				info.LicenseURL,
				info.Author,
				info.Description,
				info.Copyright,
				info.PackageURL,
				info.GitHubURL,
				info.RepositoryType,
			}
			for j, val := range row {
				cell := fmt.Sprintf("%s%d", string(rune('A'+j)), i+2)
				f.SetCellValue(sheetName, cell, val)
			}
		} else {
			info = getNPMMetadata(&pkg)
			row := []interface{}{
				info.Name + "@" + info.Version,
				info.License,
				info.Repository,
				info.LicenseURL,
				info.Author,
				info.Description,
				info.Copyright,
				info.GitHubURL,
				info.ModuleNameNoVer,
				info.Version,
			}
			for j, val := range row {
				cell := fmt.Sprintf("%s%d", string(rune('A'+j)), i+2)
				f.SetCellValue(sheetName, cell, val)
			}
		}
	}

	// Save the Excel file
	if err := f.SaveAs(outName); err != nil {
		zenity.Error("Failed to save Excel file: "+err.Error(), zenity.Title("Error"), zenity.ErrorIcon)
		return
	}

	dlg.Complete()
	zenity.Info("License report generated: "+outName, zenity.Title("Success"), zenity.InfoIcon)
}
