# src/parsers/controller_parser.py
import re
import os
from src.models import ApiEndpoint, ApiParam


def parse_controller(filepath: str) -> tuple:
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    class_name = os.path.splitext(os.path.basename(filepath))[0]

    base_path = ""
    req_map = re.search(r'@RequestMapping\s*\(\s*"([^"]+)"', content)
    if req_map:
        base_path = req_map.group(1)

    endpoints = []
    jsp_map = {}

    methods = _extract_method_blocks(content)

    for method_block in methods:
        endpoint = _parse_method_block(method_block, base_path, class_name)
        if endpoint:
            endpoints.append(endpoint)
            jsp = _extract_jsp_view(method_block)
            if jsp:
                jsp_map[endpoint.url] = jsp

    return endpoints, jsp_map


def _extract_method_blocks(content: str) -> list:
    # Find method-level mapping annotation positions (exclude class-level @RequestMapping)
    mapping_pattern = r'@(?:Get|Post|Put|Delete)Mapping'
    positions = [m.start() for m in re.finditer(mapping_pattern, content)]

    blocks = []
    for pos in positions:
        remainder = content[pos:]
        # Find the method body by looking for 'public' then the opening brace
        # This avoids confusion with {id} in URL paths
        public_match = re.search(r'public\s+', remainder)
        if not public_match:
            continue
        # Find opening brace after 'public'
        after_public = remainder[public_match.end():]
        open_brace = after_public.find('{')
        if open_brace == -1:
            continue
        open_brace += public_match.end()  # Absolute position in remainder
        # Count braces from opening brace
        depth = 0
        close_brace = -1
        for j in range(open_brace, len(remainder)):
            if remainder[j] == '{':
                depth += 1
            elif remainder[j] == '}':
                depth -= 1
                if depth == 0:
                    close_brace = j
                    break
        if close_brace != -1:
            blocks.append(remainder[:close_brace + 1])
    return blocks


def _parse_method_block(block: str, base_path: str, class_name: str) -> ApiEndpoint:
    method = "GET"
    if '@PostMapping' in block:
        method = "POST"
    elif '@PutMapping' in block:
        method = "PUT"
    elif '@DeleteMapping' in block:
        method = "DELETE"

    url = ""
    for mapping_type in ['GetMapping', 'PostMapping', 'PutMapping', 'DeleteMapping', 'RequestMapping']:
        url_match = re.search(rf'@{mapping_type}\s*\(\s*"([^"]+)"', block)
        if url_match:
            url = url_match.group(1)
            break

    full_url = _join_paths(base_path, url)

    params = []
    # @RequestParam
    for param_match in re.finditer(
        r'@RequestParam\s*\(([^)]*)\)\s*(\w+(?:<\w+>)?)\s+(\w+)',
        block
    ):
        props = param_match.group(1)
        param_type = param_match.group(2)
        param_name = param_match.group(3)
        required = True
        if 'required' in props:
            req_match = re.search(r'required\s*=\s*false', props)
            if req_match:
                required = False
        params.append(ApiParam(
            name=param_name,
            location="query",
            required=required,
            param_type=param_type,
        ))

    # @PathVariable
    for path_match in re.finditer(
        r'@PathVariable\s*(?:\([^)]*\))?\s*(\w+(?:<\w+>)?)\s+(\w+)',
        block
    ):
        params.append(ApiParam(
            name=path_match.group(2),
            location="path",
            required=True,
            param_type=path_match.group(1),
        ))

    # @RequestBody
    for body_match in re.finditer(
        r'@RequestBody\s+(\w+(?:<\w+>)?)\s+(\w+)',
        block
    ):
        params.append(ApiParam(
            name=body_match.group(2),
            location="body",
            required=True,
            param_type=body_match.group(1),
        ))

    response_type = ""
    ret_match = re.search(r'public\s+(\w+(?:<\w+[,\s\w<>]*>)?)\s+\w+\s*\(', block)
    if ret_match:
        response_type = ret_match.group(1)

    return ApiEndpoint(
        method=method,
        url=full_url,
        controller=class_name,
        request_params=params,
        response_type=response_type,
    )


def _extract_jsp_view(block: str) -> str:
    m = re.search(r'return\s*"([^"]+)"', block)
    if m:
        return m.group(1)
    return ""


def _join_paths(base: str, sub: str) -> str:
    if not base:
        return sub
    if not sub:
        return base
    base = base.rstrip('/')
    sub = sub.lstrip('/')
    return f"{base}/{sub}"
