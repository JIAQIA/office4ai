"""
PPT Data Factory

生成 PPT 事件测试数据的工厂类。
"""

from __future__ import annotations

from typing import Any


class PptDataFactory:
    """
    PPT 事件测试数据工厂。

    生成符合 PPT Socket.IO 协议的测试数据。
    """

    # === Helper builders ===

    @staticmethod
    def slide_element(
        id: str,
        type: str,
        left: float = 100.0,
        top: float = 100.0,
        width: float = 200.0,
        height: float = 150.0,
        **kwargs: Any,
    ) -> dict[str, Any]:
        """构建单个幻灯片元素。"""
        element: dict[str, Any] = {
            "id": id,
            "type": type,
            "left": left,
            "top": top,
            "width": width,
            "height": height,
        }
        element.update(kwargs)
        return element

    @staticmethod
    def slide_dimensions(
        width: float = 960.0,
        height: float = 540.0,
        aspect_ratio: str = "16:9",
    ) -> dict[str, Any]:
        """构建幻灯片尺寸信息。"""
        return {
            "width": width,
            "height": height,
            "aspectRatio": aspect_ratio,
        }

    # === Content Retrieval responses ===

    def current_slide_elements_response(
        self,
        slide_index: int = 0,
        element_count: int = 3,
    ) -> dict[str, Any]:
        """生成 ppt:get:currentSlideElements 响应数据。"""
        elements = [
            self.slide_element(
                id=f"shape-{i:03d}",
                type=["TextBox", "Image", "Shape", "Table"][i % 4],
                left=50.0 + i * 100,
                top=50.0 + i * 50,
            )
            for i in range(element_count)
        ]
        return {
            "slideIndex": slide_index,
            "elements": elements,
        }

    def slide_elements_response(
        self,
        slide_index: int = 0,
        element_count: int = 3,
    ) -> dict[str, Any]:
        """生成 ppt:get:slideElements 响应数据。"""
        elements = [
            self.slide_element(
                id=f"shape-{i:03d}",
                type=["TextBox", "Image", "Shape", "Table"][i % 4],
                left=50.0 + i * 100,
                top=50.0 + i * 50,
            )
            for i in range(element_count)
        ]
        return {
            "slideIndex": slide_index,
            "elements": elements,
        }

    def slide_screenshot_response(
        self,
        format: str = "png",
    ) -> dict[str, Any]:
        """生成 ppt:get:slideScreenshot 响应数据。"""
        return {
            "base64": "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==",
            "format": format,
        }

    def slide_info_response(
        self,
        slide_count: int = 10,
        current_index: int = 0,
        with_slide_info: bool = False,
    ) -> dict[str, Any]:
        """生成 ppt:get:slideInfo 响应数据。"""
        response: dict[str, Any] = {
            "slideCount": slide_count,
            "currentSlideIndex": current_index,
            "dimensions": self.slide_dimensions(),
        }
        if with_slide_info:
            response["slideInfo"] = {
                "layout": "Title Slide",
                "elementCount": 5,
                "background": "#FFFFFF",
            }
        return response

    def slide_layouts_response(
        self,
        layout_count: int = 5,
    ) -> dict[str, Any]:
        """生成 ppt:get:slideLayouts 响应数据。"""
        layout_names = ["Title Slide", "Title and Content", "Blank", "Two Content", "Section Header"]
        layouts = [
            {
                "id": f"layout-{i:03d}",
                "name": layout_names[i % len(layout_names)],
                "type": "custom" if i >= 3 else "builtin",
                "placeholderCount": max(0, 3 - i),
            }
            for i in range(layout_count)
        ]
        return {"layouts": layouts}

    # === Content Insertion responses ===

    def insert_text_response(
        self,
        element_id: str = "shape-015",
        slide_index: int = 0,
    ) -> dict[str, Any]:
        """生成 ppt:insert:text 响应数据。"""
        return {
            "elementId": element_id,
            "slideIndex": slide_index,
        }

    def insert_image_response(
        self,
        image_id: str = "shape-025",
        slide_index: int = 0,
    ) -> dict[str, Any]:
        """生成 ppt:insert:image 响应数据。"""
        return {
            "elementId": image_id,
            "slideIndex": slide_index,
        }

    def insert_table_response(
        self,
        element_id: str = "shape-030",
        rows: int = 3,
        columns: int = 4,
    ) -> dict[str, Any]:
        """生成 ppt:insert:table 响应数据。"""
        return {
            "elementId": element_id,
            "rows": rows,
            "columns": columns,
        }

    def insert_shape_response(
        self,
        shape_id: str = "shape-020",
        slide_index: int = 0,
    ) -> dict[str, Any]:
        """生成 ppt:insert:shape 响应数据。"""
        return {
            "elementId": shape_id,
            "slideIndex": slide_index,
        }

    # === Update responses ===

    def update_text_box_response(self) -> dict[str, Any]:
        """生成 ppt:update:textBox 响应数据。"""
        return {"updatedCount": 1}

    def update_image_response(self) -> dict[str, Any]:
        """生成 ppt:update:image 响应数据。"""
        return {"updatedCount": 1}

    def update_table_cell_response(
        self,
        updated_count: int = 1,
    ) -> dict[str, Any]:
        """生成 ppt:update:tableCell 响应数据。"""
        return {"updatedCount": updated_count}

    def update_table_row_column_response(
        self,
        updated_count: int = 1,
    ) -> dict[str, Any]:
        """生成 ppt:update:tableRowColumn 响应数据。"""
        return {"updatedCount": updated_count}

    def update_table_format_response(self) -> dict[str, Any]:
        """生成 ppt:update:tableFormat 响应数据。"""
        return {"updatedCount": 1}

    def update_element_response(self) -> dict[str, Any]:
        """生成 ppt:update:element 响应数据。"""
        return {"updatedCount": 1}

    # === Delete & Layout responses ===

    def delete_element_response(
        self,
        deleted_count: int = 1,
    ) -> dict[str, Any]:
        """生成 ppt:delete:element 响应数据。"""
        return {"deletedCount": deleted_count}

    def reorder_element_response(self) -> dict[str, Any]:
        """生成 ppt:reorder:element 响应数据。"""
        return {"reordered": True}

    # === Slide management responses ===

    def add_slide_response(
        self,
        slide_index: int = 5,
        slide_id: str = "slide-006",
    ) -> dict[str, Any]:
        """生成 ppt:add:slide 响应数据。"""
        return {
            "slideIndex": slide_index,
            "slideId": slide_id,
        }

    def delete_slide_response(
        self,
        deleted_index: int = 3,
        new_count: int = 9,
    ) -> dict[str, Any]:
        """生成 ppt:delete:slide 响应数据。"""
        return {
            "deletedIndex": deleted_index,
            "newSlideCount": new_count,
        }

    def move_slide_response(
        self,
        from_index: int = 2,
        to_index: int = 5,
    ) -> dict[str, Any]:
        """生成 ppt:move:slide 响应数据。"""
        return {
            "fromIndex": from_index,
            "toIndex": to_index,
        }

    def goto_slide_response(
        self,
        current_index: int = 5,
    ) -> dict[str, Any]:
        """生成 ppt:goto:slide 响应数据。"""
        return {
            "currentSlideIndex": current_index,
        }

    # === Error response ===

    def error_response(
        self,
        code: str = "3001",
        message: str = "Error",
    ) -> dict[str, Any]:
        """生成错误响应数据。"""
        return {
            "success": False,
            "error": {
                "code": code,
                "message": message,
            },
        }
