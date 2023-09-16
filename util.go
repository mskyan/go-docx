package docx

// maybe remove this file later.

import (
	"encoding/json"
	"fmt"
	"log"
)

func PrettyPrintStruct(structRef any) {
	b, err := json.MarshalIndent(structRef, "", "  ")
	if err != nil {
		fmt.Println("error:", err)
	}
	log.Println(string(b))
}
