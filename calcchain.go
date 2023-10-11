
package excelize

import (
	"bytes"
	"encoding/xml"
	"io"
)

// calcChainReader provides a function to get the pointer to the structure
// after deserialization of xl/calcChain.xml.
func (f *File) calcChainReader() (*xlsxCalcChain, error) {
	if f.CalcChain == nil {
		f.CalcChain = new(xlsxCalcChain)
		if err := f.xmlNewDecoder(bytes.NewReader(namespaceStrictToTransitional(f.readXML(defaultXMLPathCalcChain)))).
			Decode(f.CalcChain); err != nil && err != io.EOF {
			return f.CalcChain, err
		}
	}
	return f.CalcChain, nil
}

// calcChainWriter provides a function to save xl/calcChain.xml after
// serialize structure.
func (f *File) calcChainWriter() {
	if f.CalcChain != nil && f.CalcChain.C != nil {
		output, _ := xml.Marshal(f.CalcChain)
		f.saveFileList(defaultXMLPathCalcChain, output)
	}
}

// deleteCalcChain provides a function to remove cell reference on the
// calculation chain.
func (f *File) deleteCalcChain(index int, cell string) error {
	calc, err := f.calcChainReader()
	if err != nil {
		return err
	}
	if calc != nil {
		calc.C = xlsxCalcChainCollection(calc.C).Filter(func(c xlsxCalcChainC) bool {
			return !((c.I == index && c.R == cell) || (c.I == index && cell == "") || (c.I == 0 && c.R == cell))
		})
	}
	if len(calc.C) == 0 {
		f.CalcChain = nil
		f.Pkg.Delete(defaultXMLPathCalcChain)
		content, err := f.contentTypesReader()
		if err != nil {
			return err
		}
		content.mu.Lock()
		defer content.mu.Unlock()
		for k, v := range content.Overrides {
			if v.PartName == "/xl/calcChain.xml" {
				content.Overrides = append(content.Overrides[:k], content.Overrides[k+1:]...)
			}
		}
	}
	return err
}

type xlsxCalcChainCollection []xlsxCalcChainC

// Filter provides a function to filter calculation chain.
func (c xlsxCalcChainCollection) Filter(fn func(v xlsxCalcChainC) bool) []xlsxCalcChainC {
	var results []xlsxCalcChainC
	for _, v := range c {
		if fn(v) {
			results = append(results, v)
		}
	}
	return results
}