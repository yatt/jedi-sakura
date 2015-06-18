#! /usr/bin/python2.7
# coding: utf-8

# ソースコードを標準出力から読む

import sys
import jedi

# patch
input_path = sys.argv[1]
input_editing_path = sys.argv[2]
output_path = sys.argv[3]
cursor_line = int( sys.argv[4] )
cursor_column = int( sys.argv[5] )

# io redirect
sys.stdin = open(input_editing_path, 'r')
sys.stdout = open(output_path, 'w')

s_source = sys.stdin.read()

s_encoding = None
l_lines = s_source.splitlines()
n_lines = len(l_lines)
n_col = len(l_lines[-1])

script = jedi.Script(source = s_source, line = cursor_line, column = cursor_column, path = input_path)

completions = script.completions()
for cmpl in completions:
    if cmpl.type == 'function': # and method
        params = ', '.join(p.description for p in cmpl.params)
        print '%s(%s)' % (cmpl.name, params)
    else:
        print cmpl.name

    
