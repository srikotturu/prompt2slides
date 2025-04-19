import os
import logging
import asyncio
import argparse
import sys
from datetime import datetime
from typing import Dict, List, Any, Optional
from rich.console import Console
from rich.panel import Panel
from rich.prompt import Prompt
from rich.logging import RichHandler
from rich.markdown import Markdown
from rich.progress import Progress
from langchain_google_genai import ChatGoogleGenerativeAI

from langchain_mcp_adapters.client import MultiServerMCPClient
from langgraph.prebuilt import create_react_agent

from dotenv import load_dotenv 

load_dotenv()

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format="%(message)s",
    datefmt="[%X]",
    handlers=[RichHandler(rich_tracebacks=True)]
)
logger = logging.getLogger("ppt_assistant")

# Rich console for pretty output
console = Console()


class PowerPointAssistant:
    """PowerPoint creation assistant using LangChain and MCP servers"""
    
    def __init__(self, api_key: Optional[str] = None, verbose: bool = False):
        """Initialize the PowerPoint assistant"""
        self.api_key = api_key or os.getenv("GOOGLE_API_KEY")
        if not self.api_key:
            console.print("[bold red]Error:[/bold red] GOOGLE_API_KEY environment variable not set")
            sys.exit(1)
            
        self.verbose = verbose
        self.conversation_history = []
        # self.base_path = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
        self.session_started = datetime.now()
        
        # Configure MCP servers
        self.mcp_servers = {
            "powerpoint": {
                "command": "python",
                "args": ["mcp_server.py"],
                "transport": "stdio",
            }
        }
        
        # Initialize model
        self.model = ChatGoogleGenerativeAI(
            model="gemini-2.0-flash-lite",
            temperature=0.7,
            google_api_key=self.api_key
        )
        
        console.print(Panel.fit(
            "[bold green]PowerPoint Assistant[/bold green]\n"
            "Ask me to create presentations, add slides, format content, and more!",
            title="Welcome"
        ))
        
    def _build_system_message(self) -> str:
        """Build the system message for the assistant"""
        current_date = datetime.now().strftime("%Y-%m-%d")
        return f"""You are a PowerPoint creation assistant. Today is {current_date}.

        You can help users:
        1. Create new presentations from scratch
        2. Add slides with text, bullet points, images, and charts
        3. Format and style slide content
        4. Save presentations to files

        Always respond with clear, step-by-step explanations of what you're doing.
        When a presentation is created or modified, summarize the changes and current state.
        """

    async def _create_agent(self, client):
        """Create the React agent with tools"""
        return create_react_agent(self.model, client.get_tools())
        
    def _add_to_history(self, role: str, content: str):
        """Add a message to conversation history"""
        self.conversation_history.append({"role": role, "content": content})
        
    async def process_request(self, user_input: str) -> str:
        """Process a user request with the MCP server"""
        self._add_to_history("user", user_input)
        
        with Progress() as progress:
            task = progress.add_task("[cyan]Processing request...", total=100)
            progress.update(task, advance=30)
            
            try:
                async with MultiServerMCPClient(self.mcp_servers) as client:
                    if self.verbose:
                        logger.info("Connected to PowerPoint MCP server")
                    
                    progress.update(task, advance=30)
                    
                    # Create and invoke agent
                    agent = await self._create_agent(client)
                    response = await agent.ainvoke({
                        "messages": [
                            {"role": "system", "content": self._build_system_message()},
                            *self.conversation_history
                        ]
                    })
                    
                    progress.update(task, advance=40)
                    
                    # Extract response content
                    if isinstance(response, dict) and "messages" in response:
                        result = response["messages"][-1]["content"]
                    else:
                        result = str(response)
                    
                    self._add_to_history("assistant", result)
                    return result
            except Exception as e:
                logger.error(f"Error processing request: {str(e)}")
                return f"[bold red]Error:[/bold red] {str(e)}"
    
    def save_conversation(self, filename: str = None):
        """Save the conversation history to a file"""
        if not filename:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"ppt_conversation_{timestamp}.txt"
        
        with open(filename, 'w') as f:
            f.write(f"PowerPoint Assistant Conversation - {self.session_started}\n\n")
            for msg in self.conversation_history:
                f.write(f"{msg['role'].upper()}:\n{msg['content']}\n\n")
        
        console.print(f"[green]Conversation saved to[/green] {filename}")
        
    async def run_interactive(self):
        """Run the assistant in interactive mode"""
        console.print("\nType your requests below. Use [bold cyan]/help[/bold cyan] for commands or [bold cyan]/exit[/bold cyan] to quit.\n")
        
        while True:
            user_input = Prompt.ask("[bold green]You[/bold green]")
            
            # Handle special commands
            if user_input.lower() == "/exit":
                console.print("[yellow]Exiting PowerPoint Assistant. Goodbye![/yellow]")
                break
            elif user_input.lower() == "/help":
                self._show_help()
                continue
            elif user_input.lower().startswith("/save"):
                parts = user_input.split(maxsplit=1)
                filename = parts[1] if len(parts) > 1 else None
                self.save_conversation(filename)
                continue
            elif user_input.lower() == "/clear":
                self.conversation_history = []
                console.print("[yellow]Conversation history cleared.[/yellow]")
                continue
                
            # Process normal request
            response = await self.process_request(user_input)
            
            # Display the response
            console.print("\n[bold blue]Assistant:[/bold blue]")
            console.print(Markdown(response))
            console.print("\n" + "-" * 80 + "\n")
    
    def _show_help(self):
        """Show help information"""
        help_text = """
# PowerPoint Assistant Commands

- `/help` - Show this help message
- `/exit` - Exit the application
- `/save [filename]` - Save conversation history to a file
- `/clear` - Clear conversation history

# Example Requests

- "Create a presentation about renewable energy with 5 slides"
- "Add a slide with a comparison chart of solar vs wind energy"
- "Format all titles to be blue and centered"
- "Save the presentation as renewable_energy.pptx"
        """
        console.print(Markdown(help_text))


async def main():
    """Main entry point for the application"""
    parser = argparse.ArgumentParser(description="PowerPoint Terminal Assistant")
    parser.add_argument("--api-key", help="Google API Key (or set GOOGLE_API_KEY env var)")
    parser.add_argument("--verbose", action="store_true", help="Enable verbose logging")
    args = parser.parse_args()
    
    assistant = PowerPointAssistant(api_key=args.api_key, verbose=args.verbose)
    await assistant.run_interactive()


if __name__ == "__main__":
    try:
        asyncio.run(main())
    except KeyboardInterrupt:
        console.print("\n[yellow]Program interrupted. Exiting...[/yellow]")
        sys.exit(0)