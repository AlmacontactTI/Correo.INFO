package main

import (
	"errors"
	"fmt"
	"net/url"
	"path/filepath"
	"time"

	"fyne.io/fyne"
	"fyne.io/fyne/app"
	"fyne.io/fyne/container"
	"fyne.io/fyne/dialog"
	"fyne.io/fyne/widget"
	"github.com/tebeka/selenium"
	"github.com/tebeka/selenium/chrome"
	"github.com/xuri/excelize/v2"
)

var selectedFilePath string

func main() {
	core := app.New()
	//core.Settings().SetTheme(theme.LightTheme())
	window := core.NewWindow("Correo")
	window.Resize(fyne.NewSize(600, 500))

	titulo := widget.NewLabel("Creacion masiva correos .INFO")
	titulocenter := container.NewCenter(titulo)

	space1 := widget.NewLabel("")

	line1 := widget.NewSeparator()

	correo := widget.NewLabel("Ingrese correo ADMIN")
	correocenter := container.NewCenter(correo)
	correoentry := widget.NewEntry()

	password := widget.NewLabel("Ingrese contraseña")
	passwordcenter := container.NewCenter(password)
	passwordentry := widget.NewPasswordEntry()

	line2 := widget.NewSeparator()
	space2 := widget.NewLabel("")

	// Etiqueta para mostrar el nombre del archivo seleccionado
	selectedFileLabel := widget.NewLabel("")

	fileButton := widget.NewButton("Agregar Archivo", func() {
		openfile := dialog.NewFileOpen(func(file fyne.URIReadCloser, _ error) {
			if file == nil || file.URI() == nil {
				return
			}
			// Obtener el URI como una cadena
			uriString := file.URI().String()

			// Convertir el URI a una cadena y extraer la ruta del archivo
			uri, err := url.Parse(uriString)
			if err != nil {
				fmt.Println("Error al analizar el URI:", err)
				return
			}
			// Obtener la ruta del archivo desde el URI
			selectedFilePath = uri.Path

			// Imprimir la ruta del archivo para verificar
			fmt.Println("Ruta del archivo:", selectedFilePath)

			// Actualizar la etiqueta para mostrar el nombre del archivo seleccionado
			fileName := filepath.Base(selectedFilePath)
			selectedFileName := fileName
			selectedFileLabel.SetText(selectedFileName)
		}, window)
		openfile.Show()
	})

	crearbutton := widget.NewButton("CREAR", func() {
		correo := correoentry.Text
		contraseña := passwordentry.Text

		if correo == "" {
			dialog.ShowError(errors.New("ingrese su correo"), window)
			return
		}

		if contraseña == "" {
			dialog.ShowError(errors.New("ingrese su contraseña"), window)
			return
		}

		if selectedFilePath == "" {
			dialog.ShowError(errors.New("por favor selecciona un archivo"), window)
			return
		}

		done := make(chan struct{}) // Canal para comunicar el estado de finalización

		// Llamar a la función scrap con el correo, la contraseña y la ruta del archivo
		go scrap(correo, contraseña, selectedFilePath, done)

		// Escuchar el canal de finalización
		go func() {
			<-done // Esperar hasta que la señal de finalización sea recibida
			dialog.ShowInformation("Finalizado", "La creación masiva de correos ha finalizado", window)
		}()
	})

	contenido := container.NewVBox(titulocenter, space1, line1, correocenter, correoentry, passwordcenter, passwordentry, line2, space2, selectedFileLabel, fileButton, crearbutton)

	centeredContent := container.NewCenter(contenido)

	window.SetContent(centeredContent)
	window.ShowAndRun()
}

func scrap(correo, contraseña, filePath string, done chan struct{}) {
	defer close(done) // Cerrar el canal al finalizar la función
	// Iniciar el servicio del controlador Chrome
	servicio, err := selenium.NewChromeDriverService("./chromedriver", 4444)
	if err != nil {
		panic(err)
	}
	defer servicio.Stop()

	// Configurar las capacidades del navegador
	caps := selenium.Capabilities{}
	caps.AddChrome(chrome.Capabilities{Args: []string{
		//"--headless"
		"--start-maximized",
	}})

	// Iniciar el navegador Chrome
	driver, err := selenium.NewRemote(caps, "")
	if err != nil {
		panic(err)
	}

	// Abrir la página web
	err = driver.Get("http://webadmin.almacontactcol.info/Mondo/lang/sys/login.aspx")
	if err != nil {
		panic(err)
	}
	time.Sleep(2 * time.Second)

	//Buscar campo Email
	email, err := driver.FindElement(selenium.ByCSSSelector, "#txtUsername")
	if err != nil {
		panic(err)
	}
	//Enviar correo
	err = email.SendKeys(correo)
	if err != nil {
		panic(err)
	}
	time.Sleep(1 * time.Second)

	//Buscar campo Contraseña
	pass, err := driver.FindElement(selenium.ByCSSSelector, "#txtPassword")
	if err != nil {
		panic(err)
	}
	//Enviar contraseña
	err = pass.SendKeys(contraseña)
	if err != nil {
		panic(err)
	}
	time.Sleep(1 * time.Second)

	//Buscar Botón Login
	login, err := driver.FindElement(selenium.ByCSSSelector, "#btnLogin_js")
	if err != nil {
		panic(err)
	}
	login.Click()
	time.Sleep(2 * time.Second)

	//Buscar Cuentas
	acount, err := driver.FindElement(selenium.ByCSSSelector, "#nav_mailboxes")
	if err != nil {
		panic(err)
	}
	acount.Click()
	time.Sleep(22 * time.Second)

	// Abrir el archivo Excel
	f, err := excelize.OpenFile(filePath)
	if err != nil {
		panic(err)
	}

	// Obtener todas las filas de la hoja de cálculo
	rows, err := f.GetRows("Hoja1")
	if err != nil {
		panic(err)
	}

	// Iterar sobre cada fila del archivo Excel
	for i, row := range rows {
		if i == 0 {
			// Saltar la primera fila
			continue
		}

		// Buscar el iframe que contiene el botón "Agregar nuevo"
		iframe, err := driver.FindElement(selenium.ByCSSSelector, "#frmUserList")
		if err != nil {
			panic(err)
		}
		// Cambiar al contexto del iframe
		err = driver.SwitchFrame(iframe)
		if err != nil {
			panic(err)
		}
		// Buscar el botón "Agregar nuevo" dentro del iframe
		new, err := driver.FindElement(selenium.ByID, "btnAdd")
		if err != nil {
			panic(err)
		}
		new.Click()
		time.Sleep(1 * time.Second)
		// Después de hacer clic en "Agregar nuevo", cambiar al control de la nueva ventana/pestaña
		handles, err := driver.WindowHandles()
		if err != nil {
			panic(err)
		}
		// La nueva ventana/pestaña debería estar en la última posición
		nuevaVentana := handles[len(handles)-1]
		// Cambiar al control de la nueva ventana/pestaña
		err = driver.SwitchWindow(nuevaVentana)
		if err != nil {
			panic(err)
		}
		time.Sleep(2 * time.Second)

		// Buscar campo User
		user, err := driver.FindElement(selenium.ByCSSSelector, "#txtMailbox")
		if err != nil {
			panic(err)
		}
		// Enviar usuario
		err = user.Clear()
		if err != nil {
			panic(err)
		}
		err = user.SendKeys(row[0])
		if err != nil {
			panic(err)
		}
		time.Sleep(1 * time.Second)

		//Buscar campo contraseña
		pass2, err := driver.FindElement(selenium.ByCSSSelector, "#txtPassword")
		if err != nil {
			panic(err)
		}
		//Enviar contraseña
		err = pass2.Clear()
		if err != nil {
			panic(err)
		}
		err = pass2.SendKeys(row[2])
		if err != nil {
			panic(err)
		}
		time.Sleep(1 * time.Second)

		//Buscar campo Nombre a mostrar
		nomb, err := driver.FindElement(selenium.ByCSSSelector, "#txtDisplayName")
		if err != nil {
			panic(err)
		}
		//Enviar datos
		err = nomb.Clear()
		if err != nil {
			panic(err)
		}
		err = nomb.SendKeys(row[1])
		if err != nil {
			panic(err)
		}
		time.Sleep(1 * time.Second)

		//Buscar campo Cupo
		cupo, err := driver.FindElement(selenium.ByCSSSelector, "#txtMailboxSize")
		if err != nil {
			panic(err)
		}
		// Limpiar el contenido del campo Cupo
		err = cupo.Clear()
		if err != nil {
			panic(err)
		}
		//Enviar cupo
		err = cupo.SendKeys(row[3])
		if err != nil {
			panic(err)
		}
		time.Sleep(1 * time.Second)

		// Buscar el elemento input de tipo hidden
		inputHidden, err := driver.FindElement(selenium.ByCSSSelector, "#tbl1 > div > label > input[type=hidden]:nth-child(3)")
		if err != nil {
			panic(err)
		}
		// Obtener el valor del atributo "value"
		value, err := inputHidden.GetAttribute("value")
		if err != nil {
			panic(err)
		}
		time.Sleep(1 * time.Second)

		row = append(row, value)

		// Escribir la fila en el archivo Excel original
		for j, value := range row {
			colName, err := excelize.ColumnNumberToName(j + 1)
			if err != nil {
				panic(err)
			}
			f.SetCellValue("Hoja1", colName+fmt.Sprint(i+1), value)
		}
		// Guardar los cambios en el archivo Excel
		if err := f.Save(); err != nil {
			panic(err)
		}

		//Buscar Boton Agregar
		add, err := driver.FindElement(selenium.ByCSSSelector, "#btnAdd")
		if err != nil {
			panic(err)
		}
		add.Click()
		time.Sleep(2 * time.Second)

		alertPresent := false
		if alert, err := driver.AlertText(); err == nil {
			// Si hay una alerta presente, capturar su texto
			alertPresent = true

			f.SetCellValue("Hoja1", "F"+fmt.Sprint(i+1), alert)
			if err := f.Save(); err != nil {
				panic(err)
			}
			driver.AcceptAlert()
		}

		// Si no hay una alerta presente, simplemente continuar con el resto del código
		if !alertPresent {
			// Cambiar de vuelta al control de la ventana/pestaña original
			err = driver.SwitchWindow(handles[0])
			if err != nil {
				panic(err)
			}
		}

		add.SendKeys(selenium.ControlKey + "w")
		time.Sleep(2 * time.Second)

		// Cambiar de vuelta al control de la ventana/pestaña original
		err = driver.SwitchWindow(handles[0])
		if err != nil {
			panic(err)
		}
	}

	// Después de interactuar con los elementos dentro del iframe, es importante volver al contexto predeterminado fuera del iframe
	err = driver.SwitchFrame(nil)
	if err != nil {
		panic(err)
	}

	time.Sleep(10 * time.Second)

}
