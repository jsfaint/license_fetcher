package main

import (
	"encoding/json"
	"fmt"
	"net/http"
	"os"
	"path/filepath"
	"strings"

	"github.com/antchfx/htmlquery"
	"github.com/ncruces/zenity"
	"golang.org/x/mod/modfile"
)

type PackageInfo struct {
	Name             string
	Version          string
	License          string
	LicenseURL       string
	Author           string
	Description      string
	Copyright        string
	PackageURL       string
	GitHubURL        string
	RepositoryType   string
	Repository       string
	ModuleNameNoVer  string
}

// Package represents a dependency
type Package struct {
	Path       string
	Version    string
	GoMod      bool
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
	moduleName := file.Module.Mod.Path
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

	return packages, packageJSON.Name, nil
}

// Get metadata from pkg.go.dev
func getGoModMetadata(pkg *Package) PackageInfo {
	info := PackageInfo{
		Name:           pkg.Path,
		Version:        pkg.Version,
		PackageURL:     pkg.Path + "/@v/" + pkg.Version + ".info",
		RepositoryType: "go",
	}

	// Get license and other info from pkg.go.dev
	doc, err := htmlquery.LoadURL("https://pkg.go.dev/" + pkg.Path)
	if err == nil {
		// Find license
		node := htmlquery.FindOne(doc, `//span[contains(@class, "License")]/a`)
		if node == nil {
			node = htmlquery.FindOne(doc, `//a[contains(@href, "licenses")]`)
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
		if node != nil {
			info.Description = strings.TrimSpace(htmlquery.InnerText(node))
		}

		// Find repository link (GitHub or other)
		node = htmlquery.FindOne(doc, `//a[contains(@href, "github.com") or contains(@href, "gitlab") or contains(@href, "bitbucket")]`)
		if node != nil {
			info.GitHubURL = htmlquery.SelectAttr(node, "href")
		}

		// Find author/maintainer info
		node = htmlquery.FindOne(doc, `//span[contains(@class, "author") or contains(text(), "Author") or contains(text(), "Maintainer")]`)
		if node == nil {
			node = htmlquery.FindOne(doc, `//div[contains(@class, "metadata")]//span[contains(@class, "text-muted")]/following-sibling::span`)
		}
		if node != nil {
			info.Author = strings.TrimSpace(htmlquery.InnerText(node))
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

	// Get info from npm registry
	resp, err := http.Get("https://registry.npmjs.org/" + pkg.Path + "/" + version)
	if err == nil && resp.StatusCode == 200 {
		defer resp.Body.Close()
		var npmPkg struct {
			License  string `json:"license"`
			Licenses []struct {
				Type string `json:"type"`
			} `json:"licenses"`
			Author any `json:"author"`
			Description string `json:"description"`
			Repository struct {
				Type string `json:"type"`
				URL  string `json:"url"`
			} `json:"repository"`
			Homepage string `json:"homepage"`
		}

		if err := json.NewDecoder(resp.Body).Decode(&npmPkg); err == nil {
			if npmPkg.License != "" {
				info.License = npmPkg.License
				info.LicenseURL = "https://licenses.nuget.org/" + npmPkg.License
			} else if len(npmPkg.Licenses) > 0 {
				info.License = npmPkg.Licenses[0].Type
				info.LicenseURL = "https://licenses.nuget.org/" + npmPkg.Licenses[0].Type
			}

			if author, ok := npmPkg.Author.(map[string]any); ok {
				if name, ok := author["name"]; ok {
					info.Author = name.(string)
				}
			}

			info.Description = npmPkg.Description

			if npmPkg.Repository.URL != "" {
				info.Repository = npmPkg.Repository.URL
				info.GitHubURL = npmPkg.Repository.URL
			} else if npmPkg.Homepage != "" {
				info.Repository = npmPkg.Homepage
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
		zenity.Error("Failed to select file: "+err.Error(), zenity.Title("Error"), zenity.ErrorIcon)
		return
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

	outName := moduleName + "_license.csv"

	// create csv file
	out, err := os.Create(outName)
	if err != nil {
		zenity.Error("Failed to create output file: "+err.Error(), zenity.Title("Error"), zenity.ErrorIcon)
		return
	}
	defer out.Close()

	dlg, err := zenity.Progress(
		zenity.Title("Running..."))
	if err != nil {
		zenity.Error("Create progress dialog failed: "+err.Error(), zenity.Title("Error"), zenity.ErrorIcon)
		os.Exit(1)
	}
	defer dlg.Close()

	// Write csv header based on file type
	if isGoMod {
		fmt.Fprintln(out, "Name,License,PackageVersion,LicenseURL,Author,Description,Copyright,PackageURL,GitHubURL,RepositoryType")
	} else {
		fmt.Fprintln(out, "Module Name,License,Repository,License URL,Author,Description,Copyright,GitHub URL,Module Name (No Version),Version")
	}

	total := len(packages)
	for i, pkg := range packages {
		dlg.Value(int(float64(i) / float64(total) * 100))
		dlg.Text("Processing " + pkg.Path + "...")

		var info PackageInfo
		if isGoMod {
			info = getGoModMetadata(&pkg)
			fmt.Fprintf(out, "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n",
				escapeCSV(info.Name),
				escapeCSV(info.License),
				escapeCSV(info.Version),
				escapeCSV(info.LicenseURL),
				escapeCSV(info.Author),
				escapeCSV(info.Description),
				escapeCSV(info.Copyright),
				escapeCSV(info.PackageURL),
				escapeCSV(info.GitHubURL),
				escapeCSV(info.RepositoryType))
		} else {
			info = getNPMMetadata(&pkg)
			fmt.Fprintf(out, "%s,%s,%s,%s,%s,%s,%s,%s,%s,%s\n",
				escapeCSV(info.Name+"@"+info.Version),
				escapeCSV(info.License),
				escapeCSV(info.Repository),
				escapeCSV(info.LicenseURL),
				escapeCSV(info.Author),
				escapeCSV(info.Description),
				escapeCSV(info.Copyright),
				escapeCSV(info.GitHubURL),
				escapeCSV(info.ModuleNameNoVer),
				escapeCSV(info.Version))
		}
	}

	dlg.Complete()
	zenity.Info("License report generated: "+outName, zenity.Title("Success"), zenity.InfoIcon)
}

// Escape CSV fields that contain commas, quotes, or newlines
func escapeCSV(value string) string {
	if value == "" {
		return ""
	}
	if strings.ContainsAny(value, `",`) {
		return `"` + strings.ReplaceAll(value, `"`, `""`) + `"`
	}
	return value
}
