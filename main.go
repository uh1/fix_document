package main

import (
	"archive/zip"
	"bufio"
	"strings"
	"fmt"
	"io"
	"log"
	"os"
)

func on_err_abort(err error) {
	if err != nil {
		log.Fatal(err)
	}
}

func fix_document(filename_in string, filename_out string) {
	zr, err := zip.OpenReader(filename_in)
	on_err_abort(err)
	defer zr.Close()

	fo, err := os.Create(filename_out)
	on_err_abort(err)
	defer fo.Close()

	zw := zip.NewWriter(fo)
	defer zw.Close()

	for _, f := range zr.File {
		if f.Name == "word/document.xml" {
			fr, err := f.Open()
			on_err_abort(err)
			br := bufio.NewReader(fr)

			fw, err := zw.Create(f.Name)
			on_err_abort(err)
			bw := bufio.NewWriter(fw)

			//io.Copy(fw, rc)
			for b, err := br.ReadByte(); err != io.EOF; b, err = br.ReadByte() {
				on_err_abort(err)
				if b == 0 {continue}
				err = bw.WriteByte(b)
				on_err_abort(err)
			}
			err = bw.Flush()
			on_err_abort(err)
		} else {
			err = zw.Copy(f)
			on_err_abort(err)
		}
	}
}

func main(){
	if len(os.Args) <= 1 {
		fmt.Fprintf(os.Stderr, "Usage : %s filename.docx\n", os.Args[0])
		os.Exit(1)
	}
	filename_in := os.Args[1]
	if !strings.HasSuffix(filename_in, ".docx") {
		log.Fatal("The file does not have a docx extension")
	}
	filename_out := strings.TrimSuffix(filename_in, ".docx") + "_fixed.docx"

	fmt.Printf("Trying to fix %s into %s ...\n", filename_in, filename_out)

	fix_document(filename_in, "testdata/broken_fixed.docx")

	fmt.Println("Done")
}
