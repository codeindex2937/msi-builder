package main

import (
	"fmt"
	"io"
	"log"
	"os"
	"path/filepath"
	"strings"

	helper "github.com/codeindex2937/msi-helper"
	model "github.com/codeindex2937/msi-helper/model"
)

var builtinScriptName = "BIN_BUILTIN"

func main() {
	log.SetFlags(log.Lshortfile)
	msiName := "template.msi"
	if err := copyFile(msiName, "resource/template.msi"); err != nil {
		log.Fatal(err)
	}
	db, err := model.Open(msiName, model.MSIDBOPEN_TRANSACT)
	if err != nil {
		log.Fatal(err)
	}
	defer db.Close()

	upgradeCode := "{7325E7C4-20B5-4E5F-9B1B-0A11D6EAC8F5}"
	productVersion := "1.0.0"
	author := "author"
	updateSummary(db, author)

	if err := db.Update(model.Property{
		Property: "Manufacturer",
		Value:    author,
	}, nil, nil); err != nil {
		log.Fatal(err)
	}
	if err := db.Insert(model.Property{
		Property: "UpgradeCode",
		Value:    upgradeCode,
	}, model.Property{
		Property: "ProductName",
		Value:    "Example Product",
	}, model.Property{
		Property: "ProductVersion",
		Value:    productVersion,
	}, model.Property{
		Property: "ProductCode",
		Value:    model.NewUUID(),
	}); err != nil {
		log.Fatal(err)
	}

	installdir := model.Directory{
		Directory:  "INSTALLDIR",
		DefaultDir: "Example",
	}
	rootDir := model.Directory{
		Directory:  "TARGETDIR",
		DefaultDir: "SourceDir",
		Dirs: []*model.Directory{{
			Directory:  "ProgramFilesFolder",
			DefaultDir: ".",
			Dirs:       []*model.Directory{&installdir},
		}},
	}

	dirs := helper.SerializeDirectories(&rootDir)
	for _, dir := range dirs {
		if err := db.Insert(*dir); err != nil {
			log.Fatal(err)
		}
	}

	feat := model.Feature{
		Feature:   "main",
		Title:     model.OptString("main title"),
		Display:   model.OptInt16(2),
		Level:     1,
		Directory: model.OptString(installdir.Directory),
	}
	if err := db.Insert(feat); err != nil {
		log.Fatal(err)
	}

	h := helper.New(db, upgradeCode)

	if err := h.BlockOldVersion(productVersion); err != nil {
		log.Fatal(err)
	}

	serviceCompId := "{7325E7C4-20B5-4E5F-9B1B-0A11D6EAC8F6}"
	comp := addServiceComponent(&h, feat, serviceCompId, "resource/service", "service.exe", installdir.Directory)
	absPath, err := filepath.Abs("resource/builtin.vbs")
	if err != nil {
		log.Fatal(err)
	}
	script := helper.Script{
		Binary: model.Binary{
			Name: builtinScriptName,
			Data: model.Stream(absPath),
		},
		Type: model.CUSTOM_ACTION_TYPE_VBSCRIPT,
	}

	enableShortcut(h, script, comp, "service.exe", false)

	if err := h.LaunchAppPostInstall(script, comp, "service.exe"); err != nil {
		log.Fatal(err)
	}

	if err := h.PackFiles("cabinet.cab", "cabinet", comp); err != nil {
		log.Fatal(err)
	}

	if err := db.Commit(); err != nil {
		log.Fatal(err)
	}
}

func addServiceComponent(h *helper.Helper, feat model.Feature, compId, srcdir, keyfile, dir string) model.Component {
	comp := helper.NewComponent(feat.Feature, compId, "service", dir)

	if err := h.AddComponent(&comp, srcdir, keyfile); err != nil {
		log.Fatal(err)
	}

	if err := h.AddService("ExampleService", "example service", comp); err != nil {
		log.Fatal(err)
	}

	return comp
}

func enableShortcut(h helper.Helper, script helper.Script, comp model.Component, targetfile string, useScript bool) {
	shortcutName := "example"
	menuFolderName := "Example"

	absPath, err := filepath.Abs("resource/icon.ico")
	if err != nil {
		log.Fatal(err)
	}
	icon := model.Icon{
		Name: "icon.ico",
		Data: model.Stream(absPath),
	}
	if err := h.DB.Insert(icon); err != nil {
		log.Fatal(err)
	}

	if err := h.AddDesktopShortcut(shortcutName, model.NewUUID(), comp, &icon); err != nil {
		log.Fatal(err)
	}
	if err := h.AddMenuShortcut(shortcutName, menuFolderName, model.NewUUID(), comp, &icon); err != nil {
		log.Fatal(err)
	}

	if useScript {
		if err := h.DB.Delete(model.InstallExecuteSequence{
			Action: "CreateShortcuts",
		}, nil); err != nil {
			log.Fatal(err)
		}

		actions := []helper.Action{{
			Name:      "CreateMenuShortcut",
			Method:    "CreateShortcut",
			Sequence:  4502,
			Defered:   true,
			Condition: model.OptString(fmt.Sprintf("NOT NO_SHORTCUT AND ((%v) OR (STARTMENU_SHORTCUT_EXIST AND UPGRADE_FOUND))", model.CONDITION_POST_CLEAN_INSTALL)),
			Parameter: fmt.Sprintf("[ProgramMenuFolder]%v\n%v\n[%v]/%v\n[%v]", menuFolderName, shortcutName, comp.Directory, targetfile, comp.Directory),
		}, {
			Name:      "CreateDesktopShortcut",
			Method:    "CreateShortcut",
			Sequence:  4504,
			Defered:   true,
			Condition: model.OptString(fmt.Sprintf("NOT NO_SHORTCUT AND ((%v) OR (DESKTOP_SHORTCUT_EXIST AND UPGRADE_FOUND))", model.CONDITION_POST_CLEAN_INSTALL)),
			Parameter: fmt.Sprintf("[DesktopFolder]\n%v\n[%v]/%v\n[%v]", shortcutName, comp.Directory, targetfile, comp.Directory),
		}}
		if err := h.AddScript(script, actions); err != nil {
			log.Fatal(err)
		}
	}
}

func updateSummary(db model.Database, author string) {
	summaryInfo, err := db.OpenSummaryInformation(20)
	if err != nil {
		log.Fatal(err)
	}
	defer summaryInfo.Close()

	productName := "prod"
	title := "title"
	description := "desc"
	keywords := "key"
	arch := "x86"
	supportedLangs := []string{}
	langs := strings.Join(supportedLangs, ",")

	summaryInfo.SetProperty(model.PID_SUBJECT, productName)
	summaryInfo.SetProperty(model.PID_AUTHOR, author)
	summaryInfo.SetProperty(model.PID_TITLE, title)
	summaryInfo.SetProperty(model.PID_COMMENTS, description)
	summaryInfo.SetProperty(model.PID_KEYWORDS, keywords)

	if arch == "x64" {
		summaryInfo.SetProperty(model.PID_TEMPLATE, "x64;"+langs)
	} else {
		summaryInfo.SetProperty(model.PID_TEMPLATE, "Intel;"+langs)
	}

	summaryInfo.SetProperty(model.PID_WORDCOUNT, 2)
	summaryInfo.SetProperty(model.PID_PAGECOUNT, 200)
	summaryInfo.SetProperty(model.PID_REVNUMBER, model.NewUUID())
	if err := summaryInfo.Persist(); err != nil {
		log.Fatal(err)
	}
}

func copyFile(dst, src string) error {
	srcF, err := os.OpenFile(src, os.O_RDONLY, 0644)
	if err != nil {
		return err
	}
	defer srcF.Close()

	dstF, err := os.OpenFile(dst, os.O_TRUNC|os.O_CREATE|os.O_WRONLY, 0644)
	if err != nil {
		return err
	}
	defer dstF.Close()

	if _, err := io.Copy(dstF, srcF); err != nil {
		return err
	}
	return nil
}
