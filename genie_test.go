package excelizegenie

import (
	"github.com/xuri/excelize/v2"
	"testing"
)

func TestGenie_CreateSheet(t *testing.T) {
	type fields struct {
		File *excelize.File
	}
	type args struct {
		sheetName string
		datas     [][]interface{}
	}
	tests := []struct {
		name    string
		fields  fields
		args    args
		wantErr bool
	}{
		{
			name: "Create Sheet",
			fields: fields{
				File: excelize.NewFile(),
			},
			args: args{
				sheetName: "Sheet1",
				datas: [][]interface{}{
					{"a", "b"},
					{"a", "b"},
					{"a", "b"},
					{"a", "b"},
				},
			},
			wantErr: false,
		},
	}
	for _, tt := range tests {
		t.Run(tt.name, func(t *testing.T) {
			s := &Genie{
				File: tt.fields.File,
			}
			if err := s.CreateSheet(tt.args.sheetName, tt.args.datas); (err != nil) != tt.wantErr {
				t.Errorf("CreateSheet() error = %v, wantErr %v", err, tt.wantErr)
			}
		})
	}
}
