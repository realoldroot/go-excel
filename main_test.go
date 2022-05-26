package main

import (
	"fmt"
	"sort"
	"testing"
)

func TestA(t *testing.T) {
	//target := 7
	arr := []int{3, 1, 2, 5, 4}

	sort.Slice(arr, func(i, j int) bool {
		return arr[i] < arr[j]
	})

	var m map[int]int
	m = make(map[int]int)

	for i, v := range arr {
		m[i] = v
		fmt.Println(v)
	}
}
