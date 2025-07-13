import os
import json
import time
import numpy as np
import networkx as nx
import scipy
import pandas
import math
from openpyxl.styles.borders import Border, Side
from openpyxl.styles import Font, Alignment, NamedStyle
from openpyxl import Workbook, load_workbook
from openpyxl.drawing.image import Image
from datetime import date
import requests as RQS


def time_to_process(func):
    # Decorator for timing function

    def wrap(*args, **kwargs):
        start = time.time()
        result = func(*args, **kwargs)
        end = time.time()

        print(func.__name__, end-start)
        return result
    return wrap


@time_to_process
def prep_graph_numpy(input_array):
    nodes = np.unique(input_array)  # mapping node name --> index
    noidx = {n: i for i, n in enumerate(nodes)}  # mapping node index --> name

    n = nodes.size  # number of nodes

    numdata = np.vectorize(noidx.get)(input_array)  # replace node id by node index

    a = np.zeros((n, n))
    for tail, head in numdata:
        a[tail, head] = 1
        # A[head, tail] = 1  # add this line for undirected graph

    return a


@time_to_process
def prep_graph_networkx(input_array):
    g = nx.Graph([e for e in input_array])
    return nx.pagerank(g)


@time_to_process
def hello_world():
    print("Hello, world!")
    return


if __name__ == '__main__':
    #hello_world()
    data = np.array([['A', 'B'],
                     ['A', 'C'],
                     ['B', 'D'],
                     ['B', 'E'],
                     ['C', 'F'],
                     ['D', 'F'],
                     ['E', 'F']])
    print(prep_graph_numpy(data))
    print(prep_graph_networkx(data))

    # matrix_a = np.random.randint(1, 20, size=(5, 5), dtype=int)
    # matrix_b = np.random.randint(1, 20, size=(2, 2), dtype=int)

    # print(matrix_a)
    # print(matrix_b)
    # print("**********DOT PRODUCT***********")
    # print(matrix_a @ matrix_b)  # DOT product
    # print("*********CROSS PRODUCT**********")
    # print(np.cross(matrix_a, matrix_b))
    # print("*********MULTIPLICATION*********")
    # print(np.matmul(matrix_a, matrix_b))
