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

	"github.com/BurntSushi/toml"
	"github.com/antchfx/htmlquery"
	"github.com/ncruces/zenity"
	"github.com/xuri/excelize/v2"
	"golang.org/x/mod/modfile"
)

// createHTTPClient creates a standardized HTTP client with timeout settings
func createHTTPClient() *http.Client {
	return &http.Client{
		Timeout: 10 * time.Second,
		Transport: &http.Transport{
			MaxIdleConns:          10,
			IdleConnTimeout:       30 * time.Second,
			DisableCompression:    false,
			DisableKeepAlives:     false,
			ResponseHeaderTimeout: 5 * time.Second,
		},
	}
}

// cleanVersionString removes comparison operators and cleans up version strings
func cleanVersionString(version string) string {
	version = strings.TrimSpace(version)
	version = strings.TrimPrefix(version, ">=")
	version = strings.TrimPrefix(version, "==")
	version = strings.TrimPrefix(version, ">")
	version = strings.TrimPrefix(version, "<=")
	version = strings.TrimPrefix(version, "<")
	version = strings.TrimPrefix(version, "~=")
	version = strings.TrimPrefix(version, "^")
	version = strings.TrimPrefix(version, "~")
	version = strings.Split(version, ",")[0] // Take first part if multiple constraints
	version = strings.Split(version, " ")[0] // Take first part if space separated
	return version
}

// standardizeLicense converts various license formats to standard SPDX identifiers
func standardizeLicense(licenseName string) string {
	// Clean up common license abbreviations and variations
	switch licenseName {
	case "Apache Software License":
		return "Apache-2.0"
	case "BSD License":
		return "BSD-3-Clause"
	case "MIT License":
		return "MIT"
	case "Mozilla Public License 2.0 (MPL 2.0)":
		return "MPL-2.0"
	case "GNU General Public License v3 (GPLv3)":
		return "GPL-3.0"
	case "GNU General Public License v2 (GPLv2)":
		return "GPL-2.0"
	case "GNU Lesser General Public License v3 (LGPLv3)":
		return "LGPL-3.0"
	case "GNU Lesser General Public License v2 (LGPLv2)":
		return "LGPL-2.0"
	default:
		// Try to match common patterns
		if strings.Contains(licenseName, "Apache") {
			return "Apache-2.0"
		} else if strings.Contains(licenseName, "MIT") {
			return "MIT"
		} else if strings.Contains(licenseName, "BSD") {
			return "BSD-3-Clause"
		} else if strings.Contains(licenseName, "GPL") && strings.Contains(licenseName, "3") {
			return "GPL-3.0"
		} else if strings.Contains(licenseName, "GPL") && strings.Contains(licenseName, "2") {
			return "GPL-2.0"
		}
		return licenseName
	}
}

// extractGitHubLink extracts GitHub repository link from various sources
func extractGitHubLink(projectURLs map[string]string, homepage string) (string, string) {
	var repository, githubURL string

	// Check project URLs for GitHub link
	for key, url := range projectURLs {
		if strings.Contains(strings.ToLower(url), "github") {
			githubURL = url
		}
		// Also check for common repository keys
		if strings.Contains(strings.ToLower(key), "source") ||
			strings.Contains(strings.ToLower(key), "repository") {
			repository = url
		}
	}

	// Use homepage if no repository found
	if repository == "" && homepage != "" {
		repository = homepage
	}

	// If GitHub URL not found but repository has GitHub, use it
	if githubURL == "" && strings.Contains(strings.ToLower(repository), "github") {
		githubURL = repository
	}

	return repository, githubURL
}

// setCopyrightFromLicense sets copyright information based on license
func setCopyrightFromLicense(license string) string {
	if license != "" {
		return license + " Copyright"
	}
	return ""
}

// findLatestVersion finds the latest version from releases map
func findLatestVersion(releases map[string][]struct {
	PythonVersion string `json:"python_version"`
	UploadTime    string `json:"upload_time"`
}) string {
	latestVersion := ""
	latestTime := ""
	for ver, releaseList := range releases {
		if len(releaseList) > 0 {
			uploadTime := releaseList[0].UploadTime
			if uploadTime > latestTime {
				latestVersion = ver
				latestTime = uploadTime
			}
		}
	}
	return latestVersion
}

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
	Path      string
	Version   string
	GoMod     bool
	PyProject bool
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

// Parse pyproject.toml file
func parsePyProjectToml(filename string) ([]Package, string, error) {
	data, err := os.ReadFile(filename)
	if err != nil {
		return nil, "", err
	}

	var pyProject struct {
		Project struct {
			Name         string   `toml:"name"`
			Dependencies []string `toml:"dependencies"`
		} `toml:"project"`
		Tool struct {
			Poetry struct {
				Name            string            `toml:"name"`
				Dependencies    map[string]string `toml:"dependencies"`
				DevDependencies map[string]string `toml:"dev-dependencies"`
			} `toml:"poetry"`
		} `toml:"tool"`
		BuildSystem struct {
			Requires []string `toml:"requires"`
		} `toml:"build-system"`
	}

	if err := toml.Unmarshal(data, &pyProject); err != nil {
		return nil, "", err
	}

	var packages []Package

	// Handle Poetry dependencies
	if pyProject.Tool.Poetry.Dependencies != nil {
		for name, version := range pyProject.Tool.Poetry.Dependencies {
			// Skip poetry itself and special entries
			if name == "python" || strings.Contains(name, "poetry") {
				continue
			}
			packages = append(packages, Package{
				Path:      name,
				Version:   version,
				GoMod:     false,
				PyProject: true,
			})
		}
	}

	// Handle Poetry dev-dependencies
	if pyProject.Tool.Poetry.DevDependencies != nil {
		for name, version := range pyProject.Tool.Poetry.DevDependencies {
			// Skip poetry itself and special entries
			if name == "python" || strings.Contains(name, "poetry") {
				continue
			}
			packages = append(packages, Package{
				Path:      name,
				Version:   version,
				GoMod:     false,
				PyProject: true,
			})
		}
	}

	// Handle PEP 621 dependencies (project.dependencies)
	if len(pyProject.Project.Dependencies) > 0 {
		for _, dep := range pyProject.Project.Dependencies {
			// Parse dependency string like "requests>=2.0.0" or "numpy==1.19.0"
			parts := strings.Fields(dep)
			if len(parts) > 0 {
				name := parts[0]
				version := ""
				if len(parts) > 1 {
					version = strings.Join(parts[1:], " ")
				}
				packages = append(packages, Package{
					Path:      name,
					Version:   version,
					GoMod:     false,
					PyProject: true,
				})
			}
		}
	}

	// Determine project name
	projectName := "python-project"
	if pyProject.Tool.Poetry.Name != "" {
		projectName = pyProject.Tool.Poetry.Name
	} else if pyProject.Project.Name != "" {
		projectName = pyProject.Project.Name
	}

	return packages, projectName + "-py", nil
}

// Get metadata from PyPI
func getPyPI_Metadata(pkg *Package) PackageInfo {
	info := PackageInfo{
		Name:            pkg.Path,
		Version:         pkg.Version,
		ModuleNameNoVer: pkg.Path,
		RepositoryType:  "pypi",
	}

	// Clean version string - remove comparison operators
	version := cleanVersionString(pkg.Version)

	// Create HTTP client with timeout
	client := createHTTPClient()

	// Get info from PyPI API with context
	ctx, cancel := context.WithTimeout(context.Background(), 10*time.Second)
	defer cancel()

	// First try to get package info
	reqURL := "https://pypi.org/pypi/" + pkg.Path + "/json"
	req, err := http.NewRequestWithContext(ctx, "GET", reqURL, nil)
	if err != nil {
		return info
	}

	resp, err := client.Do(req)
	if err != nil || resp.StatusCode != 200 {
		return info
	}
	defer resp.Body.Close()

	var pypiPkg struct {
		Info struct {
			Author       string            `json:"author"`
			AuthorEmail  string            `json:"author_email"`
			Classifiers  []string          `json:"classifiers"`
			Description  string            `json:"description"`
			Summary      string            `json:"summary"`
			Home_page    string            `json:"home_page"`
			License      string            `json:"license"`
			Project_urls map[string]string `json:"project_urls"`
		} `json:"info"`
		Releases map[string][]struct {
			PythonVersion string `json:"python_version"`
			UploadTime    string `json:"upload_time"`
		} `json:"releases"`
		URLs []struct {
			Packagetype string `json:"packagetype"`
			URL         string `json:"url"`
		} `json:"urls"`
	}

	if err := json.NewDecoder(resp.Body).Decode(&pypiPkg); err == nil {
		// First, look for license in classifiers (more reliable)
		for _, classifier := range pypiPkg.Info.Classifiers {
			if strings.HasPrefix(classifier, "License :: ") {
				parts := strings.Split(classifier, " :: ")
				if len(parts) >= 3 {
					// Extract the license name (last part)
					licenseName := parts[len(parts)-1]
					info.License = standardizeLicense(licenseName)
					info.LicenseURL = "https://licenses.nuget.org/" + info.License
					break
				}
			}
		}

		// If no license found in classifiers, try license field
		if info.License == "" && pypiPkg.Info.License != "" {
			info.License = standardizeLicense(pypiPkg.Info.License)
			info.LicenseURL = "https://licenses.nuget.org/" + info.License
		}

		// Get author
		if pypiPkg.Info.Author != "" {
			info.Author = pypiPkg.Info.Author
		} else if pypiPkg.Info.AuthorEmail != "" {
			info.Author = pypiPkg.Info.AuthorEmail
		}

		// Get description
		if pypiPkg.Info.Summary != "" {
			info.Description = pypiPkg.Info.Summary
		} else if pypiPkg.Info.Description != "" {
			info.Description = pypiPkg.Info.Description
		}

		// Get repository URL
		if pypiPkg.Info.Home_page != "" {
			info.Repository = pypiPkg.Info.Home_page
			info.GitHubURL = pypiPkg.Info.Home_page
		}

		// Extract GitHub and repository links from project URLs
		repository, githubURL := extractGitHubLink(pypiPkg.Info.Project_urls, pypiPkg.Info.Home_page)
		if repository != "" {
			info.Repository = repository
		}
		if githubURL != "" {
			info.GitHubURL = githubURL
		}

		// Set copyright if we have license
		info.Copyright = setCopyrightFromLicense(info.License)

		// Try to find the latest version if we don't have a specific one
		if version == "" && len(pypiPkg.Releases) > 0 {
			latestVersion := findLatestVersion(pypiPkg.Releases)
			if latestVersion != "" {
				info.Version = latestVersion
			}
		} else if version != "" {
			info.Version = version
		}
	}

	return info
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
	client := createHTTPClient()

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

		// Set copyright from license
		info.Copyright = setCopyrightFromLicense(info.License)

		// If no license found, look for copyright mentions
		if info.License == "" {
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
	version := cleanVersionString(pkg.Version)

	// Create HTTP client with timeout
	client := createHTTPClient()

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

			// Set copyright from license
			info.Copyright = setCopyrightFromLicense(info.License)

			// If no license found, try to extract from README
			if info.License == "" && npmPkg.Readme != "" {
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
				Name:     "All Supported Format",
				Patterns: []string{"go.mod", "package.json", "pyproject.toml"},
				CaseFold: false,
			},
			{
				Name:     "Go Module",
				Patterns: []string{"go.mod"},
				CaseFold: false,
			},
			{
				Name:     "Package JSON",
				Patterns: []string{"package.json"},
				CaseFold: false,
			},
			{
				Name:     "Python Project",
				Patterns: []string{"pyproject.toml"},
				CaseFold: false,
			},
		},
	)
	if err != nil {
		// User cancelled - exit process instead of showing error dialog
		os.Exit(1)
	}

	isGoMod := strings.HasSuffix(inName, "go.mod")
	isPyProject := strings.HasSuffix(inName, "pyproject.toml")
	var moduleName string
	var packages []Package

	// Parse file
	if isGoMod {
		packages, moduleName, err = parseGoMod(inName)
	} else if isPyProject {
		packages, moduleName, err = parsePyProjectToml(inName)
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
	} else if isPyProject {
		header = []string{"Package Name", "License", "Version", "License URL", "Author", "Description", "Copyright", "Repository", "GitHub URL", "Repository Type"}
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
		} else if isPyProject {
			info = getPyPI_Metadata(&pkg)
			row := []interface{}{
				info.Name,
				info.License,
				info.Version,
				info.LicenseURL,
				info.Author,
				info.Description,
				info.Copyright,
				info.Repository,
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
