from __future__ import annotations

import math
from collections import Counter, defaultdict
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable

import ezdxf
from ezdxf.addons.importer import Importer


NETWORK_KEYWORDS = (
    "sp",
    "sprinkler",
    "헤드",
    "하향식",
    "상향식",
    "배관",
    "가지관",
    "교차",
    "주배관",
    "후렉",
    "후렉시블",
    "av",
    "alarm",
    "valve",
    "밸브",
    "알람",
    "직류티",
    "분기티",
    "엘보",
    "tee",
    "elbow",
)
NEGATIVE_NETWORK_KEYWORDS = ("소화기", "옥내소화전", "SHEET", "TEX", "F-FUR")
HEAD_RADIUS_KEYWORDS = ("반경", "radius", "r=2.6", "2.6m")
AUXILIARY_LAYERS = {"0", "L3", "L4"}
COMMANDS = ["LINE", "RECT", "CIRCLE", "TEXT", "MOVE", "COPY", "DELETE", "DELETE_WINDOW", "LAYER", "SAVE", "SAVEAS", "NEW"]


@dataclass
class CommandResult:
    ok: bool
    message: str
    saved_path: str | None = None


@dataclass
class InsertTransform:
    offset_x: float = 0.0
    offset_y: float = 0.0
    scale_x: float = 1.0
    scale_y: float = 1.0
    rotation_deg: float = 0.0

    def apply(self, x: float, y: float) -> tuple[float, float]:
        sx = x * self.scale_x
        sy = y * self.scale_y
        radians = math.radians(self.rotation_deg)
        rx = sx * math.cos(radians) - sy * math.sin(radians)
        ry = sx * math.sin(radians) + sy * math.cos(radians)
        return rx + self.offset_x, ry + self.offset_y

    def combine(self, child: "InsertTransform") -> "InsertTransform":
        origin_x, origin_y = self.apply(child.offset_x, child.offset_y)
        return InsertTransform(
            offset_x=origin_x,
            offset_y=origin_y,
            scale_x=self.scale_x * child.scale_x,
            scale_y=self.scale_y * child.scale_y,
            rotation_deg=self.rotation_deg + child.rotation_deg,
        )


class DXFWorkspace:
    def __init__(self, storage_dir: Path) -> None:
        self.storage_dir = storage_dir
        self.storage_dir.mkdir(parents=True, exist_ok=True)
        self.current_path: Path | None = None
        self.current_filename = "untitled.dxf"
        self.current_layer = "0"
        self._payload_cache: dict[tuple[bool, bool], dict[str, Any]] = {}
        self._serialized_cache: tuple[list[dict[str, Any]], dict[str, int]] | None = None
        self._new_document()

    def _new_document(self) -> None:
        self.doc = ezdxf.new("R2010")
        self.msp = self.doc.modelspace()
        self._ensure_layer(self.current_layer)
        self._invalidate_cache()

    def _invalidate_cache(self) -> None:
        self._payload_cache = {}
        self._serialized_cache = None

    def reset(self) -> None:
        self.current_path = None
        self.current_filename = "untitled.dxf"
        self.current_layer = "0"
        self._new_document()

    def _ensure_layer(self, layer_name: str) -> None:
        if layer_name not in self.doc.layers:
            self.doc.layers.add(layer_name)

    def load_file(self, source_path: Path) -> None:
        self.doc = ezdxf.readfile(source_path)
        self.msp = self.doc.modelspace()
        self.current_path = source_path
        self.current_filename = source_path.name
        self.current_layer = "0"
        self._invalidate_cache()

    def to_payload(
        self,
        include_network_entities: bool = False,
        include_network_summary: bool = False,
    ) -> dict[str, Any]:
        cache_key = (include_network_entities, include_network_summary)
        if cache_key in self._payload_cache:
            return self._payload_cache[cache_key]

        entities, unsupported = self._serialize_entities()
        recommended_layers: list[str] = []
        display_layers: list[str] = []
        network_ids: list[str] = []
        if include_network_summary:
            recommended_layers = self._recommended_network_layers(entities)
            display_layers = [layer for layer in recommended_layers if not self._is_head_radius_layer(layer)]
            if include_network_entities:
                network_ids = self._recommended_network_entity_ids(entities, display_layers)

        payload = {
            "filename": self.current_filename,
            "currentLayer": self.current_layer,
            "layers": self._layer_summary(entities, set(recommended_layers)),
            "networkLayers": display_layers,
            "networkEntityIds": network_ids,
            "entities": entities,
            "bounds": self._compute_bounds(entities),
            "unsupported": unsupported,
            "commands": COMMANDS,
        }
        self._payload_cache[cache_key] = payload
        return payload

    def save(self, filename: str | None = None) -> Path:
        if filename:
            output_path = self.storage_dir / filename
        elif self.current_path is not None:
            output_path = self.current_path
        else:
            output_path = self.storage_dir / self.current_filename
        self.doc.saveas(output_path)
        self.current_path = output_path
        self.current_filename = output_path.name
        return output_path

    def export_filtered(
        self,
        visible_layers: Iterable[str],
        filename: str | None = None,
        visible_source_handles: Iterable[str] | None = None,
    ) -> Path:
        visible_layer_set = {layer for layer in visible_layers if layer}
        visible_handle_set = {handle for handle in (visible_source_handles or []) if handle}
        if not visible_layer_set and not visible_handle_set:
            raise ValueError("No visible layers or entities to export.")

        output_name = filename or f"{Path(self.current_filename).stem}_visible.dxf"
        output_path = self.storage_dir / output_name
        new_doc = ezdxf.new(self.doc.dxfversion)
        importer = Importer(self.doc, new_doc)
        entities_to_import = []
        for entity in self.msp:
            handle = entity.dxf.handle
            layer = entity.dxf.layer
            if handle in visible_handle_set or layer in visible_layer_set:
                entities_to_import.append(entity)
        importer.import_entities(entities_to_import, target_layout=new_doc.modelspace())
        importer.finalize()
        new_doc.saveas(output_path)
        return output_path

    def execute(self, command: str, selected_ids: Iterable[str]) -> CommandResult:
        raw = command.strip()
        if not raw:
            return CommandResult(False, "Empty command.")
        parts = raw.split()
        op = parts[0].upper()
        args = parts[1:]

        try:
            if op == "NEW":
                self.reset()
                return CommandResult(True, "New document created.")
            if op == "LINE" and len(args) >= 4:
                x1, y1, x2, y2 = map(float, args[:4])
                self.msp.add_line((x1, y1), (x2, y2), dxfattribs={"layer": self.current_layer})
                self._invalidate_cache()
                return CommandResult(True, "Line created.")
            if op == "RECT" and len(args) >= 4:
                x1, y1, x2, y2 = map(float, args[:4])
                points = [(x1, y1), (x2, y1), (x2, y2), (x1, y2)]
                self.msp.add_lwpolyline(points, close=True, dxfattribs={"layer": self.current_layer})
                self._invalidate_cache()
                return CommandResult(True, "Rectangle created.")
            if op == "CIRCLE" and len(args) >= 3:
                x, y, r = map(float, args[:3])
                self.msp.add_circle((x, y), r, dxfattribs={"layer": self.current_layer})
                self._invalidate_cache()
                return CommandResult(True, "Circle created.")
            if op == "TEXT" and len(args) >= 3:
                x, y = map(float, args[:2])
                text = " ".join(args[2:])
                self.msp.add_text(text, dxfattribs={"layer": self.current_layer, "height": 200}).set_placement((x, y))
                self._invalidate_cache()
                return CommandResult(True, "Text created.")
            if op == "LAYER" and args:
                self.current_layer = " ".join(args)
                self._ensure_layer(self.current_layer)
                self._invalidate_cache()
                return CommandResult(True, f"Current layer set to {self.current_layer}.")
            if op == "SAVE":
                path = self.save()
                return CommandResult(True, f"Saved to {path.name}.", str(path))
            if op == "SAVEAS" and args:
                filename = args[0]
                if not filename.lower().endswith(".dxf"):
                    filename += ".dxf"
                path = self.save(filename)
                return CommandResult(True, f"Saved to {path.name}.", str(path))
            handles = self._selected_source_handles(selected_ids)
            if op == "DELETE_WINDOW" and len(args) >= 4:
                x1, y1, x2, y2 = map(float, args[:4])
                deleted = self._delete_window(x1, y1, x2, y2)
                return CommandResult(True, f"Deleted {deleted} entities in window.") if deleted else CommandResult(False, "No entities found in window.")
            if op == "MOVE" and len(args) >= 2:
                dx, dy = map(float, args[:2])
                moved = self._move_entities(handles, dx, dy)
                return CommandResult(True, f"Moved {moved} entities.") if moved else CommandResult(False, "No entities moved.")
            if op == "COPY" and len(args) >= 2:
                dx, dy = map(float, args[:2])
                copied = self._copy_entities(handles, dx, dy)
                return CommandResult(True, f"Copied {copied} entities.") if copied else CommandResult(False, "No entities copied.")
            if op == "DELETE":
                deleted = self._delete_entities(handles)
                return CommandResult(True, f"Deleted {deleted} entities.") if deleted else CommandResult(False, "No entities deleted.")
        except Exception as exc:  # pragma: no cover - defensive
            return CommandResult(False, f"{op} failed: {exc}")

        return CommandResult(False, "Unsupported command or invalid arguments.")

    def _selected_source_handles(self, selected_ids: Iterable[str]) -> list[str]:
        selected_set = {entity_id for entity_id in selected_ids if entity_id}
        if not selected_set:
            return []
        entities, _ = self._serialize_entities()
        handles = []
        seen = set()
        for entity in entities:
            if entity["id"] in selected_set and entity["sourceHandle"] not in seen:
                handles.append(entity["sourceHandle"])
                seen.add(entity["sourceHandle"])
        return handles

    def _move_entities(self, handles: Iterable[str], dx: float, dy: float) -> int:
        moved = 0
        for entity in self._entity_by_handles(handles):
            entity.translate(dx, dy, 0)
            moved += 1
        if moved:
            self._invalidate_cache()
        return moved

    def _copy_entities(self, handles: Iterable[str], dx: float, dy: float) -> int:
        copied = 0
        for entity in self._entity_by_handles(handles):
            clone = entity.copy()
            clone.translate(dx, dy, 0)
            self.msp.add_entity(clone)
            copied += 1
        if copied:
            self._invalidate_cache()
        return copied

    def _delete_entities(self, handles: Iterable[str]) -> int:
        deleted = 0
        for entity in list(self._entity_by_handles(handles)):
            self.msp.delete_entity(entity)
            deleted += 1
        if deleted:
            self._invalidate_cache()
        return deleted

    def _delete_window(self, x1: float, y1: float, x2: float, y2: float) -> int:
        min_x, max_x = sorted((x1, x2))
        min_y, max_y = sorted((y1, y2))
        entities, _ = self._serialize_entities()
        handles = []
        seen = set()
        for entity in entities:
            bbox = self._entity_bbox(entity)
            if bbox[2] < min_x or bbox[0] > max_x or bbox[3] < min_y or bbox[1] > max_y:
                continue
            source_handle = entity["sourceHandle"]
            if source_handle in seen:
                continue
            seen.add(source_handle)
            handles.append(source_handle)
        return self._delete_entities(handles)

    def _entity_by_handles(self, handles: Iterable[str]) -> list[Any]:
        handle_set = {handle for handle in handles if handle}
        if not handle_set:
            return []
        return [entity for entity in self.msp if entity.dxf.handle in handle_set]

    def _serialize_entities(self) -> tuple[list[dict[str, Any]], dict[str, int]]:
        if self._serialized_cache is not None:
            return self._serialized_cache
        entities: list[dict[str, Any]] = []
        unsupported: Counter[str] = Counter()
        for entity in self.msp:
            if self._skip_entity_for_projection(entity):
                continue
            serialized = self._serialize_entity(entity, entity.dxf.layer, entity.dxf.handle, InsertTransform(), None)
            if serialized:
                entities.extend(serialized)
            else:
                unsupported[entity.dxftype()] += 1
        self._serialized_cache = (entities, dict(sorted(unsupported.items())))
        return self._serialized_cache

    def _serialize_entity(
        self,
        entity: Any,
        effective_layer: str,
        source_handle: str,
        transform: InsertTransform,
        parent_insert: str | None,
    ) -> list[dict[str, Any]] | None:
        dxftype = entity.dxftype()
        handle = entity.dxf.handle
        entity_id = handle if parent_insert is None else f"{parent_insert}:{handle}"
        if self._skip_entity_for_projection(entity, effective_layer):
            return []

        if dxftype == "LINE":
            start = transform.apply(entity.dxf.start.x, entity.dxf.start.y)
            end = transform.apply(entity.dxf.end.x, entity.dxf.end.y)
            return [{
                "id": entity_id,
                "sourceHandle": source_handle,
                "layer": effective_layer,
                "type": "LINE",
                "start": {"x": start[0], "y": start[1]},
                "end": {"x": end[0], "y": end[1]},
                "color": self._entity_color(entity),
                "entityHandle": handle,
                "parentPath": parent_insert or "",
            }]

        if dxftype == "CIRCLE":
            center = transform.apply(entity.dxf.center.x, entity.dxf.center.y)
            radius = entity.dxf.radius * max(abs(transform.scale_x), abs(transform.scale_y))
            return [{
                "id": entity_id,
                "sourceHandle": source_handle,
                "layer": effective_layer,
                "type": "CIRCLE",
                "center": {"x": center[0], "y": center[1]},
                "radius": radius,
                "color": self._entity_color(entity),
                "entityHandle": handle,
                "parentPath": parent_insert or "",
            }]

        if dxftype == "ARC":
            center = transform.apply(entity.dxf.center.x, entity.dxf.center.y)
            radius = entity.dxf.radius * max(abs(transform.scale_x), abs(transform.scale_y))
            return [{
                "id": entity_id,
                "sourceHandle": source_handle,
                "layer": effective_layer,
                "type": "ARC",
                "center": {"x": center[0], "y": center[1]},
                "radius": radius,
                "startAngle": float(entity.dxf.start_angle + transform.rotation_deg),
                "endAngle": float(entity.dxf.end_angle + transform.rotation_deg),
                "color": self._entity_color(entity),
                "entityHandle": handle,
                "parentPath": parent_insert or "",
            }]

        if dxftype == "LWPOLYLINE":
            points = []
            for point in entity.get_points("xy"):
                x, y = transform.apply(point[0], point[1])
                points.append({"x": x, "y": y})
            return [{
                "id": entity_id,
                "sourceHandle": source_handle,
                "layer": effective_layer,
                "type": "LWPOLYLINE",
                "points": points,
                "closed": bool(entity.closed),
                "color": self._entity_color(entity),
                "entityHandle": handle,
                "parentPath": parent_insert or "",
            }]

        if dxftype in {"TEXT", "MTEXT"}:
            insert = transform.apply(entity.dxf.insert.x, entity.dxf.insert.y)
            text = entity.plain_text() if hasattr(entity, "plain_text") else entity.dxf.text
            height = float(getattr(entity.dxf, "char_height", getattr(entity.dxf, "height", 200.0)))
            return [{
                "id": entity_id,
                "sourceHandle": source_handle,
                "layer": effective_layer,
                "type": "TEXT",
                "insert": {"x": insert[0], "y": insert[1]},
                "text": text,
                "height": height,
                "color": self._entity_color(entity),
                "entityHandle": handle,
                "parentPath": parent_insert or "",
            }]

        if dxftype == "INSERT":
            return self._serialize_insert(entity, transform)

        return None

    def _serialize_insert(self, entity: Any, transform: InsertTransform) -> list[dict[str, Any]]:
        block_name = entity.dxf.name
        if block_name not in self.doc.blocks:
            return []

        insert_transform = InsertTransform(
            offset_x=entity.dxf.insert.x,
            offset_y=entity.dxf.insert.y,
            scale_x=float(getattr(entity.dxf, "xscale", 1.0) or 1.0),
            scale_y=float(getattr(entity.dxf, "yscale", 1.0) or 1.0),
            rotation_deg=float(getattr(entity.dxf, "rotation", 0.0) or 0.0),
        )
        combined = transform.combine(insert_transform)
        source_handle = entity.dxf.handle
        effective_layer = entity.dxf.layer
        attribute_map = {
            str(attrib.dxf.tag or "").strip(): attrib.plain_text() if hasattr(attrib, "plain_text") else str(getattr(attrib.dxf, "text", ""))
            for attrib in getattr(entity, "attribs", [])
        }
        items: list[dict[str, Any]] = []
        for child in self.doc.blocks[block_name]:
            child_layer = child.dxf.layer
            child_effective_layer = effective_layer if child_layer == "0" else child_layer
            serialized = self._serialize_entity(
                child,
                child_effective_layer,
                source_handle,
                combined,
                parent_insert=source_handle,
            )
            if serialized:
                for item in serialized:
                    item["blockName"] = block_name
                    item["blockAttributes"] = attribute_map
                    item["insertPoint"] = {"x": entity.dxf.insert.x, "y": entity.dxf.insert.y}
                items.extend(serialized)
        return items

    def _recommended_network_layers(self, entities: list[dict[str, Any]]) -> list[str]:
        groups = self._group_entities_by_source(entities)
        components = self._build_network_components(groups)
        selected_layers: set[str] = set()
        for component in components:
            if not self._component_is_network(component):
                continue
            for group in component["groups"]:
                if group["pipeScore"] > 0 or group["seedScore"] > 0:
                    selected_layers.update(
                        layer for layer in group["layers"]
                        if self._is_network_layer(layer) and not self._is_head_radius_layer(layer)
                    )
        if selected_layers:
            return sorted(selected_layers)

        layer_counts: Counter[str] = Counter(entity["layer"] for entity in entities)
        return sorted(
            layer for layer in layer_counts
            if self._is_network_layer(layer) and not self._is_head_radius_layer(layer)
        )

    def _recommended_network_entity_ids(
        self,
        entities: list[dict[str, Any]],
        display_layers: list[str],
    ) -> list[str]:
        if not entities or not display_layers:
            return []
        display_layer_set = set(display_layers)
        groups = self._group_entities_by_source(entities)
        components = self._build_network_components(groups)
        ids: list[str] = []
        for component in components:
            if not self._component_is_network(component):
                continue
            for group in component["groups"]:
                if not self._group_should_survive(group, component, display_layer_set):
                    continue
                ids.extend(
                    entity["id"]
                    for entity in group["entities"]
                    if not self._is_head_radius_entity(entity)
                )
        return sorted(set(ids))

    def _build_network_components(self, groups: list[dict[str, Any]]) -> list[dict[str, Any]]:
        prepared = [self._prepare_group_for_network(group) for group in groups if not group["radius_only"]]
        if not prepared:
            return []
        adjacency: list[set[int]] = [set() for _ in prepared]
        tolerance = 140.0
        for left_idx in range(len(prepared)):
            for right_idx in range(left_idx + 1, len(prepared)):
                if not self._bbox_close(prepared[left_idx]["bbox"], prepared[right_idx]["bbox"], 420.0):
                    continue
                if self._groups_connected_topology(prepared[left_idx], prepared[right_idx], tolerance):
                    adjacency[left_idx].add(right_idx)
                    adjacency[right_idx].add(left_idx)

        visited: set[int] = set()
        components: list[dict[str, Any]] = []
        for start_idx in range(len(prepared)):
            if start_idx in visited:
                continue
            stack = [start_idx]
            visited.add(start_idx)
            indices: list[int] = []
            while stack:
                idx = stack.pop()
                indices.append(idx)
                for nxt in adjacency[idx]:
                    if nxt not in visited:
                        visited.add(nxt)
                        stack.append(nxt)
            component_groups = [prepared[idx] for idx in indices]
            component_adjacency = {
                group["sourceHandle"]: {prepared[nxt]["sourceHandle"] for nxt in adjacency[idx] if nxt in indices}
                for idx, group in enumerate(prepared)
                if idx in indices
            }
            components.append(self._score_network_component(component_groups, component_adjacency))
        return components

    def _prepare_group_for_network(self, group: dict[str, Any]) -> dict[str, Any]:
        prepared = dict(group)
        entity_text = " ".join(self._entity_hint_text(entity) for entity in group["entities"]).lower()
        seed_score = 0
        pipe_score = 0
        if any(self._is_network_layer(layer) for layer in group["layers"]):
            pipe_score += 3
        if self._group_has_head_hint(group, entity_text):
            seed_score += 4
        if self._group_has_valve_hint(group, entity_text):
            seed_score += 4
        if self._group_has_size_hint(group, entity_text):
            pipe_score += 1
        if group["length"] >= 250.0 and group["segments"]:
            pipe_score += 2
        if group["text_only"]:
            pipe_score -= 4
        if self._group_is_annotation(group, entity_text):
            pipe_score -= 5
        if self._group_is_isolated_short_segment(group):
            pipe_score -= 2
        prepared["seedScore"] = seed_score
        prepared["pipeScore"] = pipe_score
        prepared["headLike"] = self._group_has_head_hint(group, entity_text)
        prepared["valveLike"] = self._group_has_valve_hint(group, entity_text)
        prepared["fittingLike"] = self._group_is_fitting_like(group, entity_text)
        prepared["annotationLike"] = self._group_is_annotation(group, entity_text)
        return prepared

    def _entity_hint_text(self, entity: dict[str, Any]) -> str:
        parts = [entity.get("layer", ""), entity.get("blockName", ""), entity.get("text", ""), entity.get("parentPath", "")]
        for key, value in (entity.get("blockAttributes") or {}).items():
            parts.append(str(key))
            parts.append(str(value))
        return " ".join(part for part in parts if part)

    def _group_has_head_hint(self, group: dict[str, Any], hint_text: str) -> bool:
        head_tokens = ("head", "sprink", "upright", "pendent", "sidewall", "nozzle")
        return any(token in hint_text for token in head_tokens) or (
            any(entity["type"] == "CIRCLE" and entity.get("radius", 0.0) < 500.0 for entity in group["entities"])
            and len(group["segments"]) <= 4
        )

    def _group_has_valve_hint(self, group: dict[str, Any], hint_text: str) -> bool:
        valve_tokens = ("valve", "alarm", "drain", "check", "gate", "butterfly", "av")
        return any(token in hint_text for token in valve_tokens)

    def _group_has_size_hint(self, group: dict[str, Any], hint_text: str) -> bool:
        if "dn" in hint_text:
            return True
        return any(token in hint_text for token in ("25a", "32a", "40a", "50a", "65a", "80a", "100a"))

    def _group_is_annotation(self, group: dict[str, Any], hint_text: str) -> bool:
        negative_tokens = ("dim", "anno", "note", "grid", "axis", "room", "arch", "struct", "text", "leader")
        if any(token in hint_text for token in negative_tokens):
            return True
        return group["text_only"]

    def _group_is_fitting_like(self, group: dict[str, Any], hint_text: str) -> bool:
        fitting_tokens = ("tee", "elbow", "cross", "reducer", "bend", "fitting")
        if any(token in hint_text for token in fitting_tokens):
            return True
        if self._is_network_symbol_block(group):
            return True
        return group["length"] <= 800.0 and 2 <= len(group["segments"]) <= 6 and group["span"] <= 600.0

    def _group_is_isolated_short_segment(self, group: dict[str, Any]) -> bool:
        return len(group["segments"]) == 1 and group["length"] < 120.0 and not group["anchor"]

    def _groups_connected_topology(self, left: dict[str, Any], right: dict[str, Any], tolerance: float) -> bool:
        if left["sourceHandle"] == right["sourceHandle"]:
            return True
        if left["annotationLike"] or right["annotationLike"]:
            return False
        left_endpoints = left["terminals"] or left["points"][:2]
        right_endpoints = right["terminals"] or right["points"][:2]
        if self._min_distance_between_points(left_endpoints, right_endpoints) <= tolerance:
            return True
        if left_endpoints and right["segments"] and self._min_distance_points_to_segments(left_endpoints, right["segments"]) <= tolerance:
            return True
        if right_endpoints and left["segments"] and self._min_distance_points_to_segments(right_endpoints, left["segments"]) <= tolerance:
            return True
        if self._segments_cross_without_endpoint(left["segments"], right["segments"], tolerance):
            return False
        if self._can_heal_gap(left, right):
            return True
        if (left["fittingLike"] or right["fittingLike"] or left["valveLike"] or right["valveLike"]) and self._bbox_close(left["bbox"], right["bbox"], 180.0):
            return self._linear_entities_distance(left["segments"] or [((left["bbox"][0], left["bbox"][1]), (left["bbox"][2], left["bbox"][3]))], right["segments"] or [((right["bbox"][0], right["bbox"][1]), (right["bbox"][2], right["bbox"][3]))]) <= 220.0
        return False

    def _segments_cross_without_endpoint(
        self,
        left_segments: list[tuple[tuple[float, float], tuple[float, float]]],
        right_segments: list[tuple[tuple[float, float], tuple[float, float]]],
        tolerance: float,
    ) -> bool:
        if not left_segments or not right_segments:
            return False
        for left_start, left_end in left_segments:
            for right_start, right_end in right_segments:
                if self._segment_to_segment_distance(left_start, left_end, right_start, right_end) > tolerance:
                    continue
                if self._min_distance_between_points([left_start, left_end], [right_start, right_end]) <= tolerance:
                    return False
                return True
        return False

    def _can_heal_gap(self, left: dict[str, Any], right: dict[str, Any]) -> bool:
        if not left["segments"] or not right["segments"]:
            return False
        if not (left["fittingLike"] or right["fittingLike"] or left["valveLike"] or right["valveLike"] or left["headLike"] or right["headLike"]):
            return False
        return self._linear_entities_distance(left["segments"], right["segments"]) <= 180.0

    def _score_network_component(self, groups: list[dict[str, Any]], adjacency: dict[str, set[str]]) -> dict[str, Any]:
        n_heads = sum(1 for group in groups if group["headLike"])
        n_valves = sum(1 for group in groups if group["valveLike"])
        n_fittings = sum(1 for group in groups if group["fittingLike"])
        pipe_groups = [group for group in groups if group["pipeScore"] > 0 or group["segments"]]
        continuity = sum(len(neighbors) for neighbors in adjacency.values()) / max(len(groups), 1)
        annotation_groups = sum(1 for group in groups if group["annotationLike"])
        isolated_groups = sum(1 for group in groups if len(adjacency.get(group["sourceHandle"], set())) == 0)
        score = (
            5 * n_heads
            + 3 * n_valves
            + 2 * n_fittings
            + 1 * continuity
            + 1 * len(pipe_groups)
            - 4 * annotation_groups
            - 2 * isolated_groups
        )
        return {
            "groups": groups,
            "adjacency": adjacency,
            "nHeads": n_heads,
            "nValves": n_valves,
            "nFittings": n_fittings,
            "nPipes": len(pipe_groups),
            "score": score,
        }

    def _component_is_network(self, component: dict[str, Any]) -> bool:
        if component["nPipes"] == 0:
            return False
        if component["nHeads"] == 0 and component["nValves"] == 0:
            return False
        return component["score"] >= 4

    def _group_should_survive(self, group: dict[str, Any], component: dict[str, Any], display_layer_set: set[str]) -> bool:
        if group["annotationLike"] or group["radius_only"]:
            return False
        if group["headLike"] or group["valveLike"] or group["fittingLike"]:
            return True
        if any(layer in display_layer_set for layer in group["layers"]) and group["pipeScore"] > -1:
            return True
        degree = len(component["adjacency"].get(group["sourceHandle"], set()))
        if degree >= 2 and group["segments"]:
            return True
        if group["pipeScore"] >= 2:
            return True
        return False

    def _network_candidates(
        self,
        source_groups: list[dict[str, Any]],
        display_layer_set: set[str],
        seed_points: list[tuple[float, float]],
    ) -> list[dict[str, Any]]:
        selected = []
        for group in source_groups:
            if group["display"]:
                selected.append(group)
                continue
            if self._is_auxiliary_network_entity(group, seed_points):
                selected.append(group)
        return selected

    def _is_auxiliary_network_entity(self, group: dict[str, Any], seed_points: list[tuple[float, float]]) -> bool:
        if not group["layers"] & AUXILIARY_LAYERS and not group["anchor"]:
            return False
        if group["radius_only"]:
            return False
        if not seed_points:
            return False

        distance = self._min_distance_points_to_group(seed_points, group)
        if self._is_network_symbol_block(group) and distance <= 550.0:
            return True
        if group["layers"] <= {"L3", "L4"} and group["length"] <= 180.0 and distance <= 420.0:
            return True
        if distance > 300.0:
            return False

        if group["text_only"]:
            return False
        if not self._looks_like_pipe_connector(group):
            return False
        if group["layers"] <= {"L3", "L4"} and not group["terminals"]:
            return False
        if group["span"] <= 12.0:
            return group["anchor"]
        return True

    def _looks_like_pipe_connector(self, group: dict[str, Any]) -> bool:
        type_counts = group["typeCounts"]
        allowed_types = {"LINE", "LWPOLYLINE"}
        if self._is_network_symbol_block(group):
            return True
        if any(entity_type not in allowed_types for entity_type in type_counts):
            return False
        if group["closedCount"] > 0:
            return False

        segment_count = len(group["segments"])
        if segment_count == 0:
            return False
        if segment_count > 6:
            return False

        layers = group["layers"]
        length = group["length"]
        span = max(group["span"], 1.0)
        if length < 20.0:
            return False
        tortuosity = length / span
        min_x, min_y, max_x, max_y = group["bbox"]
        width = max_x - min_x
        height = max_y - min_y

        if layers <= {"L3", "L4"}:
            if length > 420.0:
                return False
            if span > 420.0:
                return False
            if segment_count > 2:
                return False
            if tortuosity > 1.2:
                return False
            return True

        if layers == {"0"}:
            if length > 820.0:
                return False
            if span > 430.0:
                return False
            if not (3 <= segment_count <= 6):
                return False
            if not (1.15 <= tortuosity <= 2.3):
                return False
            if min(width, height) > 260.0 and tortuosity < 1.25:
                return False
            return True

        if length > 1200.0:
            return False
        if span > 900.0:
            return False
        if tortuosity > 2.4:
            return False
        if min(width, height) > 450.0 and tortuosity > 1.4:
            return False

        return True

    def _is_network_symbol_block(self, group: dict[str, Any]) -> bool:
        block_names = group.get("blockNames", set())
        if not block_names:
            return False
        if any(name in {"A$C78BA73EA", "A$C476E3A35", "A$C0B056EF9", "WTT", "A13", "A$C7822437B"} for name in block_names):
            return True
        type_counts = group["typeCounts"]
        if "TEXT" in type_counts:
            return False
        total_entities = sum(type_counts.values())
        if total_entities > 8:
            return False
        if group["span"] > 450.0 or group["length"] > 900.0:
            return False
        if any(name.startswith("A$C") for name in block_names):
            return True
        return False

    def _connected_components(
        self,
        groups: list[dict[str, Any]],
        tolerance: float,
    ) -> list[tuple[list[dict[str, Any]], dict[int, set[int]]]]:
        if not groups:
            return []

        grid_size = max(tolerance * 1.5, 150.0)
        cells: defaultdict[tuple[int, int], list[int]] = defaultdict(list)
        bboxes = [group["bbox"] for group in groups]
        adjacency: list[set[int]] = [set() for _ in groups]

        for idx, bbox in enumerate(bboxes):
            min_x, min_y, max_x, max_y = bbox
            cell_min_x = math.floor((min_x - tolerance) / grid_size)
            cell_max_x = math.floor((max_x + tolerance) / grid_size)
            cell_min_y = math.floor((min_y - tolerance) / grid_size)
            cell_max_y = math.floor((max_y + tolerance) / grid_size)
            for cell_x in range(cell_min_x, cell_max_x + 1):
                for cell_y in range(cell_min_y, cell_max_y + 1):
                    cells[(cell_x, cell_y)].append(idx)

        checked_pairs: set[tuple[int, int]] = set()
        for cell_indices in cells.values():
            for left_idx in cell_indices:
                for right_idx in cell_indices:
                    if left_idx >= right_idx:
                        continue
                    pair = (left_idx, right_idx)
                    if pair in checked_pairs:
                        continue
                    checked_pairs.add(pair)
                    left = groups[left_idx]
                    right = groups[right_idx]
                    if not self._bbox_close(bboxes[left_idx], bboxes[right_idx], tolerance):
                        continue
                    if self._groups_connected(left, right, tolerance):
                        adjacency[left_idx].add(right_idx)
                        adjacency[right_idx].add(left_idx)

        visited = set()
        components: list[tuple[list[dict[str, Any]], dict[int, set[int]]]] = []
        for start_idx in range(len(groups)):
            if start_idx in visited:
                continue
            stack = [start_idx]
            visited.add(start_idx)
            component_indices = []
            while stack:
                idx = stack.pop()
                component_indices.append(idx)
                for nxt in adjacency[idx]:
                    if nxt not in visited:
                        visited.add(nxt)
                        stack.append(nxt)
            component = [groups[idx] for idx in component_indices]
            index_map = {global_idx: local_idx for local_idx, global_idx in enumerate(component_indices)}
            component_adjacency = {
                index_map[global_idx]: {index_map[nxt] for nxt in adjacency[global_idx] if nxt in index_map}
                for global_idx in component_indices
            }
            components.append((component, component_adjacency))

        return components

    def _select_components(
        self,
        components: list[tuple[list[dict[str, Any]], dict[int, set[int]]]],
        display_layer_set: set[str],
    ) -> list[dict[str, Any]]:
        kept: list[dict[str, Any]] = []
        for component, adjacency in components:
            if not component:
                continue
            has_display_layer = any(group["display"] for group in component)
            anchor_count = sum(1 for group in component if group["anchor"])
            linear_length = sum(group["length"] for group in component)
            entity_count = sum(len(group["entities"]) for group in component)
            if not has_display_layer:
                continue
            if anchor_count == 0 and linear_length < 400.0 and entity_count < 3:
                continue
            kept.extend(self._prune_component_groups(component, adjacency))
        return kept

    def _prune_component_groups(
        self,
        component: list[dict[str, Any]],
        adjacency: dict[int, set[int]],
    ) -> list[dict[str, Any]]:
        protected = {idx for idx, group in enumerate(component) if group["display"] or group["anchor"]}
        if not protected:
            return []

        active = set(range(len(component)))
        degrees = {idx: len([nxt for nxt in adjacency.get(idx, set()) if nxt in active]) for idx in active}
        queue = [idx for idx in active if idx not in protected and degrees[idx] <= 1]

        while queue:
            idx = queue.pop()
            if idx not in active:
                continue
            active.remove(idx)
            for neighbor in adjacency.get(idx, set()):
                if neighbor not in active:
                    continue
                degrees[neighbor] = len([nxt for nxt in adjacency.get(neighbor, set()) if nxt in active])
                if neighbor not in protected and degrees[neighbor] <= 1:
                    queue.append(neighbor)

        pruned = []
        protected_groups = [component[idx] for idx in protected if idx in active]
        for idx in sorted(active):
            group = component[idx]
            if group["display"] or group["anchor"]:
                pruned.append(group)
                continue
            protected_neighbor_count = sum(1 for other in adjacency.get(idx, set()) if other in active and other in protected)
            terminal_support = self._terminal_support_count(group, protected_groups)
            distinct_terminal_groups = self._distinct_terminal_group_support(group, protected_groups)
            protected_touch_count = self._protected_touch_count(group, protected_groups)
            if group["layers"] <= {"L3", "L4"}:
                if group["length"] <= 180.0:
                    pruned.append(group)
                    continue
                if group["length"] <= 140.0 and (terminal_support >= 1 or protected_touch_count >= 1):
                    pruned.append(group)
                    continue
                if (
                    group["length"] <= 320.0
                    and (
                        distinct_terminal_groups >= 2
                        or terminal_support >= 2
                        or (terminal_support >= 1 and protected_touch_count >= 2)
                    )
                ):
                    pruned.append(group)
                continue
            if group["layers"] == {"0"}:
                if self._is_network_symbol_block(group):
                    pruned.append(group)
                    continue
                if distinct_terminal_groups >= 2 or terminal_support >= 3 or (protected_neighbor_count >= 1 and protected_touch_count >= 2):
                    pruned.append(group)
                continue
            if protected_neighbor_count >= 1 or terminal_support >= 2:
                pruned.append(group)
        return pruned

    def _terminal_support_count(self, group: dict[str, Any], protected_groups: list[dict[str, Any]]) -> int:
        if not group["terminals"]:
            return 0
        supported = 0
        for terminal in group["terminals"]:
            for other in protected_groups:
                if self._min_distance_between_points([terminal], other["points"]) <= 170.0:
                    supported += 1
                    break
        return supported

    def _distinct_terminal_group_support(self, group: dict[str, Any], protected_groups: list[dict[str, Any]]) -> int:
        if not group["terminals"]:
            return 0
        touched_handles: set[str] = set()
        for terminal in group["terminals"]:
            nearest_handle = None
            nearest_distance = math.inf
            for other in protected_groups:
                distance = self._min_distance_between_points([terminal], other["points"])
                if distance < nearest_distance and distance <= 170.0:
                    nearest_distance = distance
                    nearest_handle = other["sourceHandle"]
            if nearest_handle:
                touched_handles.add(nearest_handle)
        return len(touched_handles)

    def _protected_touch_count(self, group: dict[str, Any], protected_groups: list[dict[str, Any]]) -> int:
        touched = 0
        for other in protected_groups:
            if self._groups_connected(group, other, tolerance=140.0):
                touched += 1
        return touched

    def _is_network_layer(self, layer_name: str) -> bool:
        lowered = layer_name.lower()
        if self._is_head_radius_layer(layer_name):
            return False
        if any(keyword.lower() in lowered for keyword in NEGATIVE_NETWORK_KEYWORDS):
            return False
        return any(keyword in lowered for keyword in NETWORK_KEYWORDS)

    def _skip_entity_for_projection(self, entity: Any, layer_name: str | None = None) -> bool:
        layer_name = layer_name or entity.dxf.layer
        if not self._layer_is_projectable(layer_name):
            return True
        paperspace = int(getattr(entity.dxf, "paperspace", 0) or 0)
        if paperspace:
            return True
        if int(getattr(entity.dxf, "invisible", 0) or 0):
            return True
        return False

    def _layer_is_projectable(self, layer_name: str) -> bool:
        if layer_name not in self.doc.layers:
            return True
        layer = self.doc.layers.get(layer_name)
        try:
            if layer.is_off():
                return False
        except AttributeError:
            pass
        try:
            if layer.is_frozen():
                return False
        except AttributeError:
            pass
        try:
            if layer.is_locked():
                return False
        except AttributeError:
            pass
        flags = int(getattr(layer.dxf, "flags", 0) or 0)
        if flags & 1:
            return False
        if flags & 4:
            return False
        return True

    def _is_head_radius_layer(self, layer_name: str) -> bool:
        lowered = layer_name.lower().replace(" ", "")
        return any(keyword in lowered for keyword in HEAD_RADIUS_KEYWORDS)

    def _is_head_radius_entity(self, entity: dict[str, Any]) -> bool:
        if self._is_head_radius_layer(entity["layer"]):
            return True
        if entity["type"] == "CIRCLE" and entity.get("radius", 0.0) >= 2400.0:
            return True
        text = (entity.get("text") or "").lower().replace(" ", "")
        return any(keyword in text for keyword in HEAD_RADIUS_KEYWORDS)

    def _is_anchor_entity(self, entity: dict[str, Any]) -> bool:
        hint_text = self._entity_hint_text(entity).lower()
        if any(keyword.lower() in hint_text for keyword in NEGATIVE_NETWORK_KEYWORDS):
            return False
        if any(keyword in hint_text for keyword in NETWORK_KEYWORDS):
            return True
        return entity["type"] == "TEXT" and self._group_has_size_hint({"entities": [entity]}, hint_text)

    def _group_entities_by_source(self, entities: list[dict[str, Any]]) -> list[dict[str, Any]]:
        by_source: defaultdict[str, list[dict[str, Any]]] = defaultdict(list)
        for entity in entities:
            if self._is_head_radius_entity(entity):
                continue
            by_source[entity["sourceHandle"]].append(entity)

        groups = []
        for source_handle, source_entities in by_source.items():
            all_points = []
            all_segments = []
            layers = {entity["layer"] for entity in source_entities}
            block_names = {entity.get("blockName") for entity in source_entities if entity.get("blockName")}
            anchor = any(self._is_anchor_entity(entity) for entity in source_entities)
            display = any(self._is_network_layer(layer) and not self._is_head_radius_layer(layer) for layer in layers)
            text_only = all(entity["type"] == "TEXT" for entity in source_entities)
            radius_only = all(self._is_head_radius_entity(entity) for entity in source_entities)
            type_counts: Counter[str] = Counter(entity["type"] for entity in source_entities)
            closed_count = sum(1 for entity in source_entities if entity["type"] == "LWPOLYLINE" and entity.get("closed"))
            length = 0.0
            for entity in source_entities:
                points = self._entity_key_points(entity)
                all_points.extend(points[:12])
                all_segments.extend(self._entity_segments(entity))
                length += self._entity_linear_length(entity)
            if not all_points and all_segments:
                all_points.extend([segment[0] for segment in all_segments[:12]])
                all_points.extend([segment[1] for segment in all_segments[:12]])
            bbox = self._merge_bboxes(self._entity_bbox(entity) for entity in source_entities)
            terminals = self._group_terminal_points(all_segments, all_points)
            groups.append(
                {
                    "sourceHandle": source_handle,
                    "entities": source_entities,
                    "layers": layers,
                    "blockNames": block_names,
                    "anchor": anchor,
                    "display": display,
                    "text_only": text_only,
                    "radius_only": radius_only,
                    "typeCounts": dict(type_counts),
                    "closedCount": closed_count,
                    "length": length,
                    "span": math.hypot(bbox[2] - bbox[0], bbox[3] - bbox[1]),
                    "bbox": bbox,
                    "points": all_points[:32],
                    "segments": all_segments[:48],
                    "terminals": terminals[:12],
                }
            )
        return groups

    def _group_terminal_points(
        self,
        segments: list[tuple[tuple[float, float], tuple[float, float]]],
        points: list[tuple[float, float]],
    ) -> list[tuple[float, float]]:
        if not segments:
            return points[:4]
        degree: dict[tuple[int, int], int] = defaultdict(int)
        point_lookup: dict[tuple[int, int], tuple[float, float]] = {}
        for start, end in segments:
            for point in (start, end):
                key = (round(point[0]), round(point[1]))
                degree[key] += 1
                point_lookup[key] = point
        terminals = [point_lookup[key] for key, value in degree.items() if value == 1]
        return terminals or points[:4]

    def _candidate_source_groups(self, display_layer_set: set[str]) -> list[dict[str, Any]]:
        groups = []
        for entity in self.msp:
            layer = entity.dxf.layer
            if layer not in display_layer_set:
                continue
            serialized = self._serialize_entity(entity, layer, entity.dxf.handle, InsertTransform(), None)
            if not serialized:
                continue
            groups.extend(self._group_entities_by_source(serialized))

        seed_points = [point for group in groups for point in group["points"]]
        if not seed_points:
            return groups

        seed_grid = self._build_point_grid(seed_points, cell_size=400.0)
        for entity in self.msp:
            layer = entity.dxf.layer
            if layer not in AUXILIARY_LAYERS:
                continue
            bbox = self._source_entity_bbox(entity)
            if not self._bbox_near_seed_grid(bbox, seed_grid, cell_size=400.0, tolerance=280.0):
                continue
            serialized = self._serialize_entity(entity, layer, entity.dxf.handle, InsertTransform(), None)
            if not serialized:
                continue
            aux_groups = self._group_entities_by_source(serialized)
            for group in aux_groups:
                if self._is_auxiliary_network_entity(group, seed_points):
                    groups.append(group)
        return groups

    def _groups_connected(self, left: dict[str, Any], right: dict[str, Any], tolerance: float) -> bool:
        if left["sourceHandle"] == right["sourceHandle"]:
            return True
        if left["radius_only"] or right["radius_only"]:
            return False
        if left["segments"] and right["segments"]:
            if self._linear_entities_distance(left["segments"], right["segments"]) <= tolerance:
                return True
        if self._min_distance_between_points(left["points"], right["points"]) <= tolerance:
            return True
        if left["points"] and right["segments"]:
            if self._min_distance_points_to_segments(left["points"], right["segments"]) <= tolerance:
                return True
        if right["points"] and left["segments"]:
            if self._min_distance_points_to_segments(right["points"], left["segments"]) <= tolerance:
                return True
        return False

    def _entity_key_points(self, entity: dict[str, Any]) -> list[tuple[float, float]]:
        entity_type = entity["type"]
        if entity_type == "LINE":
            return [
                (entity["start"]["x"], entity["start"]["y"]),
                (entity["end"]["x"], entity["end"]["y"]),
            ]
        if entity_type == "LWPOLYLINE":
            return [(point["x"], point["y"]) for point in entity["points"]]
        if entity_type == "CIRCLE":
            center_x = entity["center"]["x"]
            center_y = entity["center"]["y"]
            radius = entity["radius"]
            return [
                (center_x, center_y),
                (center_x + radius, center_y),
                (center_x - radius, center_y),
                (center_x, center_y + radius),
                (center_x, center_y - radius),
            ]
        if entity_type == "ARC":
            center_x = entity["center"]["x"]
            center_y = entity["center"]["y"]
            radius = entity["radius"]
            start_radians = math.radians(entity["startAngle"])
            end_radians = math.radians(entity["endAngle"])
            return [
                (center_x, center_y),
                (center_x + radius * math.cos(start_radians), center_y + radius * math.sin(start_radians)),
                (center_x + radius * math.cos(end_radians), center_y + radius * math.sin(end_radians)),
            ]
        if entity_type == "TEXT":
            return [(entity["insert"]["x"], entity["insert"]["y"])]
        return []

    def _entity_segments(self, entity: dict[str, Any]) -> list[tuple[tuple[float, float], tuple[float, float]]]:
        if entity["type"] == "LINE":
            return [(
                (entity["start"]["x"], entity["start"]["y"]),
                (entity["end"]["x"], entity["end"]["y"]),
            )]
        if entity["type"] != "LWPOLYLINE":
            return []
        points = self._entity_key_points(entity)
        segments = [(points[idx], points[idx + 1]) for idx in range(len(points) - 1)]
        if entity.get("closed") and len(points) > 2:
            segments.append((points[-1], points[0]))
        return segments

    def _entity_bbox(self, entity: dict[str, Any]) -> tuple[float, float, float, float]:
        points = self._entity_key_points(entity)
        if not points:
            return (0.0, 0.0, 0.0, 0.0)
        xs = [point[0] for point in points]
        ys = [point[1] for point in points]
        return min(xs), min(ys), max(xs), max(ys)

    def _merge_bboxes(self, bboxes: Iterable[tuple[float, float, float, float]]) -> tuple[float, float, float, float]:
        bboxes = list(bboxes)
        if not bboxes:
            return (0.0, 0.0, 0.0, 0.0)
        return (
            min(bbox[0] for bbox in bboxes),
            min(bbox[1] for bbox in bboxes),
            max(bbox[2] for bbox in bboxes),
            max(bbox[3] for bbox in bboxes),
        )

    def _build_point_grid(
        self,
        points: list[tuple[float, float]],
        cell_size: float,
    ) -> defaultdict[tuple[int, int], list[tuple[float, float]]]:
        grid: defaultdict[tuple[int, int], list[tuple[float, float]]] = defaultdict(list)
        for x, y in points:
            grid[(math.floor(x / cell_size), math.floor(y / cell_size))].append((x, y))
        return grid

    def _bbox_near_seed_grid(
        self,
        bbox: tuple[float, float, float, float],
        seed_grid: defaultdict[tuple[int, int], list[tuple[float, float]]],
        cell_size: float,
        tolerance: float,
    ) -> bool:
        min_x, min_y, max_x, max_y = bbox
        cell_min_x = math.floor((min_x - tolerance) / cell_size)
        cell_max_x = math.floor((max_x + tolerance) / cell_size)
        cell_min_y = math.floor((min_y - tolerance) / cell_size)
        cell_max_y = math.floor((max_y + tolerance) / cell_size)
        for cell_x in range(cell_min_x, cell_max_x + 1):
            for cell_y in range(cell_min_y, cell_max_y + 1):
                if (cell_x, cell_y) in seed_grid:
                    return True
        return False

    def _source_entity_bbox(self, entity: Any) -> tuple[float, float, float, float]:
        dxftype = entity.dxftype()
        if dxftype == "LINE":
            xs = [entity.dxf.start.x, entity.dxf.end.x]
            ys = [entity.dxf.start.y, entity.dxf.end.y]
            return min(xs), min(ys), max(xs), max(ys)
        if dxftype in {"CIRCLE", "ARC"}:
            center_x = entity.dxf.center.x
            center_y = entity.dxf.center.y
            radius = entity.dxf.radius
            return center_x - radius, center_y - radius, center_x + radius, center_y + radius
        if dxftype == "LWPOLYLINE":
            points = list(entity.get_points("xy"))
            if not points:
                return (0.0, 0.0, 0.0, 0.0)
            xs = [point[0] for point in points]
            ys = [point[1] for point in points]
            return min(xs), min(ys), max(xs), max(ys)
        if dxftype in {"TEXT", "MTEXT"}:
            return entity.dxf.insert.x, entity.dxf.insert.y, entity.dxf.insert.x, entity.dxf.insert.y
        if dxftype == "INSERT":
            return entity.dxf.insert.x, entity.dxf.insert.y, entity.dxf.insert.x, entity.dxf.insert.y
        return (0.0, 0.0, 0.0, 0.0)

    def _bbox_close(
        self,
        left: tuple[float, float, float, float],
        right: tuple[float, float, float, float],
        tolerance: float,
    ) -> bool:
        return not (
            left[2] + tolerance < right[0]
            or right[2] + tolerance < left[0]
            or left[3] + tolerance < right[1]
            or right[3] + tolerance < left[1]
        )

    def _entities_connected(self, left: dict[str, Any], right: dict[str, Any], tolerance: float) -> bool:
        if left["sourceHandle"] == right["sourceHandle"]:
            return True
        if self._is_head_radius_entity(left) or self._is_head_radius_entity(right):
            return False

        left_segments = self._entity_segments(left)
        right_segments = self._entity_segments(right)
        if left_segments and right_segments:
            return self._linear_entities_distance(left_segments, right_segments) <= tolerance

        left_points = self._entity_key_points(left)
        right_points = self._entity_key_points(right)
        if self._min_distance_between_points(left_points, right_points) <= tolerance:
            return True
        if left_points and right_segments and self._min_distance_points_to_segments(left_points, right_segments) <= tolerance:
            return True
        if right_points and left_segments and self._min_distance_points_to_segments(right_points, left_segments) <= tolerance:
            return True
        return False

    def _entity_span(self, entity: dict[str, Any]) -> float:
        min_x, min_y, max_x, max_y = self._entity_bbox(entity)
        return math.hypot(max_x - min_x, max_y - min_y)

    def _entity_linear_length(self, entity: dict[str, Any]) -> float:
        if entity["type"] == "LINE":
            points = self._entity_key_points(entity)
            return math.dist(points[0], points[1])
        if entity["type"] != "LWPOLYLINE":
            return 0.0
        total = 0.0
        for start, end in self._entity_segments(entity):
            total += math.dist(start, end)
        return total

    def _min_distance_between_points(
        self,
        left_points: list[tuple[float, float]],
        right_points: list[tuple[float, float]],
    ) -> float:
        if not left_points or not right_points:
            return math.inf
        return min(math.dist(left, right) for left in left_points for right in right_points)

    def _min_distance_points_to_entity(self, points: list[tuple[float, float]], entity: dict[str, Any]) -> float:
        segments = self._entity_segments(entity)
        if segments:
            return self._min_distance_points_to_segments(points, segments)
        key_points = self._entity_key_points(entity)
        return self._min_distance_between_points(points, key_points)

    def _min_distance_points_to_group(self, points: list[tuple[float, float]], group: dict[str, Any]) -> float:
        if group["segments"]:
            return self._min_distance_points_to_segments(points, group["segments"])
        return self._min_distance_between_points(points, group["points"])

    def _min_distance_points_to_segments(
        self,
        points: list[tuple[float, float]],
        segments: list[tuple[tuple[float, float], tuple[float, float]]],
    ) -> float:
        if not points or not segments:
            return math.inf
        return min(
            self._point_to_segment_distance(point[0], point[1], start[0], start[1], end[0], end[1])
            for point in points
            for start, end in segments
        )

    def _linear_entities_distance(
        self,
        left_segments: list[tuple[tuple[float, float], tuple[float, float]]],
        right_segments: list[tuple[tuple[float, float], tuple[float, float]]],
    ) -> float:
        return min(
            self._segment_to_segment_distance(left_start, left_end, right_start, right_end)
            for left_start, left_end in left_segments
            for right_start, right_end in right_segments
        )

    def _segment_to_segment_distance(
        self,
        left_start: tuple[float, float],
        left_end: tuple[float, float],
        right_start: tuple[float, float],
        right_end: tuple[float, float],
    ) -> float:
        return min(
            self._point_to_segment_distance(left_start[0], left_start[1], right_start[0], right_start[1], right_end[0], right_end[1]),
            self._point_to_segment_distance(left_end[0], left_end[1], right_start[0], right_start[1], right_end[0], right_end[1]),
            self._point_to_segment_distance(right_start[0], right_start[1], left_start[0], left_start[1], left_end[0], left_end[1]),
            self._point_to_segment_distance(right_end[0], right_end[1], left_start[0], left_start[1], left_end[0], left_end[1]),
        )

    def _point_to_segment_distance(self, px: float, py: float, ax: float, ay: float, bx: float, by: float) -> float:
        dx = bx - ax
        dy = by - ay
        length_sq = dx * dx + dy * dy
        if length_sq == 0:
            return math.hypot(px - ax, py - ay)
        t = ((px - ax) * dx + (py - ay) * dy) / length_sq
        t = max(0.0, min(1.0, t))
        sx = ax + t * dx
        sy = ay + t * dy
        return math.hypot(px - sx, py - sy)

    def _layer_summary(self, entities: list[dict[str, Any]], recommended_layers: set[str]) -> list[dict[str, Any]]:
        counts: Counter[str] = Counter(entity["layer"] for entity in entities)
        layers = []
        for layer in sorted(entry.dxf.name for entry in self.doc.layers):
            if layer in counts or layer == self.current_layer:
                layers.append(
                    {
                        "name": layer,
                        "entityCount": counts.get(layer, 0),
                        "recommendedNetwork": layer in recommended_layers and not self._is_head_radius_layer(layer),
                    }
                )
        return layers

    def _entity_color(self, entity: Any) -> str:
        color_index = int(getattr(entity.dxf, "color", 7) or 7)
        palette = {
            1: "#ff6b6b",
            2: "#ffd166",
            3: "#06d6a0",
            4: "#4cc9f0",
            5: "#b5179e",
            6: "#f72585",
            7: "#d8dee9",
        }
        return palette.get(color_index, "#d8dee9")

    def _compute_bounds(self, entities: list[dict[str, Any]]) -> dict[str, float]:
        if not entities:
            return {"minX": -500.0, "minY": -500.0, "maxX": 500.0, "maxY": 500.0}

        min_x = math.inf
        min_y = math.inf
        max_x = -math.inf
        max_y = -math.inf
        for entity in entities:
            bbox = self._entity_bbox(entity)
            min_x = min(min_x, bbox[0])
            min_y = min(min_y, bbox[1])
            max_x = max(max_x, bbox[2])
            max_y = max(max_y, bbox[3])

        if not math.isfinite(min_x):
            return {"minX": -500.0, "minY": -500.0, "maxX": 500.0, "maxY": 500.0}

        padding = max((max_x - min_x) * 0.03, (max_y - min_y) * 0.03, 100.0)
        return {
            "minX": min_x - padding,
            "minY": min_y - padding,
            "maxX": max_x + padding,
            "maxY": max_y + padding,
        }


def build_explicit_pipe_mismatches(
    cad_path: Path | None,
    sdf_path: Path | None,
    *,
    length_tol_abs_m: float = 0.5,
    length_tol_ratio: float = 0.05,
    elevation_tol_m: float = 0.1,
) -> list[dict[str, Any]]:
    if cad_path is None or sdf_path is None or not Path(cad_path).exists() or not Path(sdf_path).exists():
        return []

    mismatches: list[dict[str, Any]] = []
    workspace = DXFWorkspace(Path(cad_path).parent / "_cad_compare_workspace")
    workspace.load_file(Path(cad_path))
    payload = workspace.to_payload(include_network_entities=True, include_network_summary=True)
    entities = payload.get("entities") or []
    network_layers = set(payload.get("networkLayers") or [])
    filtered = [e for e in entities if not network_layers or e.get("layer") in network_layers]

    cad_line_count = sum(1 for e in filtered if e.get("type") in {"LINE", "LWPOLYLINE", "ARC"})
    cad_circle_count = sum(1 for e in filtered if e.get("type") == "CIRCLE")

    import xml.etree.ElementTree as ET

    root = ET.parse(sdf_path).getroot()
    sdf_pipe_count = len(root.findall(".//Pipe"))
    sdf_nozzle_count = len(root.findall(".//Nozzle"))

    if cad_circle_count != sdf_nozzle_count:
        mismatches.append(
            {
                "type": "head_count_match",
                "message": f"head/node count mismatch: drawing {cad_circle_count}, schematic {sdf_nozzle_count}",
                "drawing_value": cad_circle_count,
                "schematic_value": sdf_nozzle_count,
            }
        )
    if cad_line_count and abs(cad_line_count - sdf_pipe_count) / max(sdf_pipe_count, 1) > max(length_tol_ratio, 0.01):
        mismatches.append(
            {
                "type": "topology_connectivity_match",
                "message": f"topology/connectivity mismatch: drawing segments {cad_line_count}, schematic pipes {sdf_pipe_count}",
                "drawing_value": cad_line_count,
                "schematic_value": sdf_pipe_count,
                "length_tol_abs_m": length_tol_abs_m,
                "length_tol_ratio": length_tol_ratio,
                "elevation_tol_m": elevation_tol_m,
            }
        )
    return mismatches
