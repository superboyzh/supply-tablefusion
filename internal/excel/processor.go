package excel

import (
	"bytes"
	"errors"
	"fmt"
	"io"

	"github.com/xuri/excelize/v2"
)

type SourceType string

const (
	SourceTypeOutbound SourceType = "outbound"
	SourceTypeWeidian  SourceType = "weidian"
)

var (
	ErrUnsupportedSourceType  = errors.New("unsupported source type")
	ErrMappingsNotImplemented = errors.New("excel field mappings are not implemented yet")
)

func (s SourceType) Valid() bool {
	return s == SourceTypeOutbound || s == SourceTypeWeidian
}

func Transform(input io.Reader, sourceType SourceType) (*bytes.Buffer, error) {
	if !sourceType.Valid() {
		return nil, fmt.Errorf("%w: %s", ErrUnsupportedSourceType, sourceType)
	}

	workbook, err := excelize.OpenReader(input)
	if err != nil {
		return nil, fmt.Errorf("open workbook: %w", err)
	}
	defer workbook.Close()

	return nil, ErrMappingsNotImplemented
}
