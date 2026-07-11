"""Dependency-free play-geometry normalization helpers."""

import math


def merge_block_caps(routes, lines):
    """Fold PowerPoint stub + cap pairs into conservative block terminals."""

    def seg_dir(pts):
        (x1, y1), (x2, y2) = pts[0], pts[-1]
        distance = math.hypot(x2 - x1, y2 - y1) or 1.0
        return (x2 - x1) / distance, (y2 - y1) / distance

    remaining = []
    for line in lines:
        points = line["points"]
        if len(points) == 2 and line.get("end", "none") == "none" and not line.get("dash"):
            (x1, y1), (x2, y2) = points
            cap_length = math.hypot(x2 - x1, y2 - y1)
            if 0.005 <= cap_length <= 0.09:
                midpoint = ((x1 + x2) / 2, (y1 + y2) / 2)
                cap_direction = seg_dir(points)
                candidates = []
                for index, route in enumerate(routes):
                    if (
                        route.get("end", "none") != "none"
                        or route.get("dash")
                        or route.get("color") != line.get("color")
                        or len(route.get("points", [])) < 2
                    ):
                        continue
                    tip = route["points"][-1]
                    tip_distance = math.hypot(midpoint[0] - tip[0], midpoint[1] - tip[1])
                    if tip_distance > 0.035:
                        continue
                    route_direction = seg_dir(route["points"][-2:])
                    perpendicularity = abs(
                        route_direction[0] * cap_direction[0]
                        + route_direction[1] * cap_direction[1]
                    )
                    if perpendicularity < 0.45:
                        candidates.append((tip_distance, perpendicularity, index))
                if candidates:
                    _distance, _perpendicularity, route_index = min(candidates)
                    routes[route_index]["end"] = "block"
                    continue
        remaining.append(line)
    return routes, remaining
