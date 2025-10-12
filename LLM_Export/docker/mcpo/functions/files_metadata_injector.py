"""
title: Files Metadata Injector
author: GlissemanTV
version: 1.0.0
description: Automatically injects file metadata so that tools can use them
"""

from typing import Optional
from pydantic import BaseModel, Field


class Filter:
    class Valves(BaseModel):
        priority: int = Field(
            default=0,
            description="Filter priority (the lower the number, the sooner it runs)",
        )
        enabled: bool = Field(
            default=True, description="Activates or deactivates the filter"
        )

    def __init__(self):
        self.valves = self.Valves()

    def inlet(
        self, body: dict, __user__: Optional[dict] = None, __event_emitter__=None
    ) -> dict:
        """
        Injects file metadata into the context so that tools can use it
        """

        if not self.valves.enabled:
            return body

        files = body.get("files", [])

        if not files:
            return body

        files_metadata = {
            "files": [{"id": f.get("id"), "name": f.get("name")} for f in files]
        }

        metadata_message = (
            f"[SYSTEM CONTEXT - Files Available]\n"
            f"The following files are available in this conversation:\n"
            f"Files count: {len(files_metadata['files'])}\n"
            f"Files list: {', '.join([f['name'] for f in files_metadata['files']])}\n\n"
            f"File metadata for tools:\n{files_metadata}\n"
            f"[END SYSTEM CONTEXT]\n\n"
            f"You can now call the appropriate tools to process these files."
        )

        messages = body.get("messages", [])

        if messages and not any(
            "[SYSTEM CONTEXT - Files Available]" in msg.get("content", "")
            for msg in messages
        ):

            messages.insert(0, {"role": "system", "content": metadata_message})

            body["messages"] = messages

        return body