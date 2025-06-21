/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

// Store conversation history for chat mode
let conversationHistory = [];
// Store hovering chat history separately
let hoveringChatHistory = [];
// Track if hovering chat is open
let isHoveringChatOpen = false;

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    
    // Set up chat toggle
    const chatToggle = document.getElementById("chat-toggle") as HTMLInputElement;
    chatToggle.addEventListener("change", toggleChatMode);
    
    // Set up chat send button
    const chatSendButton = document.getElementById("chat-send");
    if (chatSendButton) {
      chatSendButton.addEventListener("click", sendChatMessage);
    }
    
    // Set up chat input enter key
    const chatInput = document.getElementById("chat-input") as HTMLInputElement;
    if (chatInput) {
      chatInput.addEventListener("keypress", (e) => {
        if (e.key === "Enter") {
          sendChatMessage();
        }
      });
    }
    
    // Set up hovering chat button
    const hoveringChatButton = document.getElementById("hovering-chat-button");
    if (hoveringChatButton) {
      hoveringChatButton.addEventListener("click", toggleHoveringChat);
    }
    
    // Set up close hovering chat button
    const closeHoveringChatButton = document.getElementById("close-hovering-chat");
    if (closeHoveringChatButton) {
      closeHoveringChatButton.addEventListener("click", closeHoveringChat);
    }
    
    // Set up hovering chat send button
    const hoveringChatSendButton = document.getElementById("hovering-chat-send");
    if (hoveringChatSendButton) {
      hoveringChatSendButton.addEventListener("click", sendHoveringChatMessage);
    }
    
    // Set up hovering chat input enter key
    const hoveringChatInput = document.getElementById("hovering-chat-input") as HTMLInputElement;
    if (hoveringChatInput) {
      hoveringChatInput.addEventListener("keypress", (e) => {
        if (e.key === "Enter") {
          sendHoveringChatMessage();
        }
      });
    }
    
    // Initialize the UI based on chat toggle state
    toggleChatMode();
  }
});

function toggleChatMode() {
  const chatToggle = document.getElementById("chat-toggle") as HTMLInputElement;
  const chatContainer = document.getElementById("chat-container");
  const hoveringChatButton = document.getElementById("hovering-chat-button");
  
  if (chatToggle.checked) {
    // Chat mode enabled
    chatContainer.style.display = "flex";
    hoveringChatButton.style.display = "none";
    closeHoveringChat(); // Close hovering chat if open
  } else {
    // Cell mode enabled
    chatContainer.style.display = "none";
    hoveringChatButton.style.display = "flex";
  }
}

function toggleHoveringChat() {
  const hoveringChatContainer = document.getElementById("hovering-chat-container");
  
  if (isHoveringChatOpen) {
    closeHoveringChat();
  } else {
    // Open hovering chat
    hoveringChatContainer.style.display = "flex";
    hoveringChatContainer.classList.add("fade-in");
    isHoveringChatOpen = true;
    
    // Focus the input
    const hoveringChatInput = document.getElementById("hovering-chat-input") as HTMLInputElement;
    if (hoveringChatInput) {
      hoveringChatInput.focus();
    }
  }
}

function closeHoveringChat() {
  const hoveringChatContainer = document.getElementById("hovering-chat-container");
  hoveringChatContainer.style.display = "none";
  hoveringChatContainer.classList.remove("fade-in");
  isHoveringChatOpen = false;
}

async function sendHoveringChatMessage() {
  const hoveringChatInput = document.getElementById("hovering-chat-input") as HTMLInputElement;
  const message = hoveringChatInput.value.trim();
  
  if (!message) return;
  
  // Add user message to hovering chat
  addMessageToHoveringChat(message, "user");
  hoveringChatInput.value = "";
  
  // Process with AI
  const resultInfo = document.getElementById("result-info");
  resultInfo.textContent = "Processing your request...";
  
  try {
    // Check if the message is about adding a column
    if (message.toLowerCase().includes("column") || 
        message.toLowerCase().includes("add") || 
        message.toLowerCase().includes("insert")) {
      
      // Try to handle as an Excel modification
      const handled = await handleExcelModification(message);
      if (handled) {
        resultInfo.textContent = "";
        return;
      }
    }
    
    // If not handled as Excel modification, process as normal chat
    // Add message to hovering chat history
    hoveringChatHistory.push({ role: "user", content: message });
    
    // Get AI response
    const aiResponse = await sendToOpenAI(message, true, hoveringChatHistory);
    
    // Add AI response to hovering chat
    addMessageToHoveringChat(aiResponse, "ai");
    
    // Add to hovering chat history
    hoveringChatHistory.push({ role: "assistant", content: aiResponse });
    
    resultInfo.textContent = "";
  } catch (error) {
    console.error("Chat error:", error);
    resultInfo.textContent = "Error: " + (error.message || "Could not process your request");
  }
}

function addMessageToHoveringChat(message: string, sender: "user" | "ai") {
  const hoveringChatMessages = document.getElementById("hovering-chat-messages");
  const messageElement = document.createElement("div");
  messageElement.classList.add("chat-message");
  messageElement.classList.add(sender === "user" ? "user-message" : "ai-message");
  messageElement.textContent = message;
  hoveringChatMessages.appendChild(messageElement);
  
  // Scroll to bottom
  hoveringChatMessages.scrollTop = hoveringChatMessages.scrollHeight;
}

async function sendChatMessage() {
  const chatInput = document.getElementById("chat-input") as HTMLInputElement;
  const message = chatInput.value.trim();
  
  if (!message) return;
  
  // Add user message to chat
  addMessageToChat(message, "user");
  chatInput.value = "";
  
  // Process with AI
  const resultInfo = document.getElementById("result-info");
  resultInfo.textContent = "Processing your request...";
  
  try {
    // Check if the message is about adding a column
    if (message.toLowerCase().includes("column") || 
        message.toLowerCase().includes("add") || 
        message.toLowerCase().includes("insert")) {
      
      // Try to handle as an Excel modification
      const handled = await handleExcelModification(message);
      if (handled) {
        resultInfo.textContent = "";
        return;
      }
    }
    
    // If not handled as Excel modification, process as normal chat
    // Add message to conversation history
    conversationHistory.push({ role: "user", content: message });
    
    // Get AI response
    const aiResponse = await sendToOpenAI(message, true, conversationHistory);
    
    // Add AI response to chat
    addMessageToChat(aiResponse, "ai");
    
    // Add to conversation history
    conversationHistory.push({ role: "assistant", content: aiResponse });
    
    resultInfo.textContent = "";
  } catch (error) {
    console.error("Chat error:", error);
    resultInfo.textContent = "Error: " + (error.message || "Could not process your request");
  }
}

async function handleExcelModification(message: string): Promise<boolean> {
  try {
    // Ask AI to interpret the user's intent
    const interpretPrompt = `
    Analyze this request: "${message}"
    
    If this is a request to add a column to Excel, extract the column name.
    Return ONLY a JSON object in this format:
    {
      "isAddColumn": true or false,
      "columnName": "the name of the column to add if applicable"
    }
    
    Return ONLY the JSON with no additional text.`;
    
    const interpretResponse = await sendToOpenAI(interpretPrompt, false);
    
    try {
      // Try to parse the response as JSON
      const parsed = JSON.parse(interpretResponse);
      
      if (parsed.isAddColumn && parsed.columnName) {
        // It's a request to add a column
        await addColumnToExcel(parsed.columnName);
        
        // Add response to the appropriate chat interface
        const chatToggle = document.getElementById("chat-toggle") as HTMLInputElement;
        if (chatToggle.checked) {
        addMessageToChat(`I've added a "${parsed.columnName}" column to your Excel sheet.`, "ai");
        } else if (isHoveringChatOpen) {
          addMessageToHoveringChat(`I've added a "${parsed.columnName}" column to your Excel sheet.`, "ai");
        }
        
        return true;
      }
    } catch (parseError) {
      console.error("Failed to parse AI interpretation:", parseError);
    }
    
    return false;
  } catch (error) {
    console.error("Error handling Excel modification:", error);
    return false;
  }
}

async function addColumnToExcel(columnName: string) {
  try {
    // Show processing message
    const resultInfo = document.getElementById("result-info");
    resultInfo.textContent = `Adding ${columnName} column...`;
    
    await Excel.run(async (context) => {
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("values");
      range.load("rowCount");
      range.load("columnCount");
      
      await context.sync();
      
      // Get the data from the selected range
      const rangeValues = range.values;
      
      // Only proceed if there's data in the range
      if (rangeValues && rangeValues.length > 0) {
        // Generate appropriate data for the column
        const dataPrompt = `
        Generate appropriate sample data for an Excel column named "${columnName}".
        I need ${rangeValues.length} values.
        The data should be realistic for a column with this name.
        Return ONLY a JSON array with the data values and nothing else.`;
        
        const dataResponse = await sendToOpenAI(dataPrompt, false);
        
        let columnData;
        try {
          // Try to parse the AI-generated data
          columnData = JSON.parse(dataResponse);
          
          // Ensure we have the right number of items
          if (!Array.isArray(columnData) || columnData.length !== rangeValues.length) {
            throw new Error("Invalid data format");
          }
        } catch (parseError) {
          // Fallback to simple data if parsing fails
          console.error("Failed to parse AI-generated data:", parseError);
          columnData = [];
          for (let i = 0; i < rangeValues.length; i++) {
            columnData.push(`${columnName}${i+1}`);
          }
        }
        
        // Get the column to the right of the selected range
        const newColumnRange = range.getOffsetRange(0, range.columnCount);
        
        // Create column values array
        const columnValues = columnData.map(item => [item]);
        
        // Set the values
        newColumnRange.values = columnValues;
        
        await context.sync();
        resultInfo.textContent = `${columnName} column added successfully!`;
      } else {
        resultInfo.textContent = "Please select a range with data first.";
      }
    });
  } catch (error) {
    console.error(`Error adding ${columnName} column:`, error);
    const resultInfo = document.getElementById("result-info");
    resultInfo.textContent = "Error: " + (error.message || `Could not add ${columnName} column`);
  }
}

function addMessageToChat(message: string, sender: "user" | "ai") {
  const chatMessages = document.getElementById("chat-messages");
  const messageElement = document.createElement("div");
  messageElement.classList.add("chat-message");
  messageElement.classList.add(sender === "user" ? "user-message" : "ai-message");
  messageElement.textContent = message;
  chatMessages.appendChild(messageElement);
  
  // Scroll to bottom
  chatMessages.scrollTop = chatMessages.scrollHeight;
}

export async function run() {
  try {
    // Show processing message
    const resultInfo = document.getElementById("result-info");
    resultInfo.textContent = "Processing your data...";
    
    // Check if we're in chat mode
    const chatToggle = document.getElementById("chat-toggle") as HTMLInputElement;
    const isChatMode = chatToggle.checked;
    
    await Excel.run(async (context) => {
      // Get the selected range
      const range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("values");
      range.load("rowCount");
      range.load("columnCount");
      
      await context.sync();
      
      // Get the data from the selected range
      const rangeValues = range.values;
      
      // Only proceed if there's data in the range
      if (rangeValues && rangeValues.length > 0) {
        resultInfo.textContent = "Analyzing data with AI...";
        
        // Create a prompt based on the selected data
        const dataString = JSON.stringify(rangeValues);
        const prompt = `Analyze this Excel data and provide insights: ${dataString}`;
        
        try {
          // Reset conversation history when starting a new analysis
          conversationHistory = [{ role: "user", content: prompt }];
          
          // Get AI response
          const aiResponse = await sendToOpenAI(prompt, false);
          
          if (isChatMode) {
            // Add the data as a user message (simplified representation)
            addMessageToChat("I need analysis for my selected Excel data", "user");
            
            // Add AI response to chat
            addMessageToChat(aiResponse, "ai");
            
            // Add to conversation history
            conversationHistory.push({ role: "assistant", content: aiResponse });
            
            resultInfo.textContent = "";
          } else {
            // Cell mode - write to Excel
            // Create a single cell for the output rather than trying to match dimensions
            const outputCell = range.getCell(range.rowCount, 0).getOffsetRange(1, 0);
            outputCell.values = [[aiResponse]];
            
            await context.sync();
            resultInfo.textContent = "Analysis complete! Results added below your data.";
            
            // Show hovering chat with the analysis result
            if (!isHoveringChatOpen) {
              toggleHoveringChat();
            }
            
            // Clear previous messages
            const hoveringChatMessages = document.getElementById("hovering-chat-messages");
            hoveringChatMessages.innerHTML = "";
            
            // Add AI response to hovering chat
            addMessageToHoveringChat(aiResponse, "ai");
            
            // Reset hovering chat history
            hoveringChatHistory = [
              { role: "user", content: prompt },
              { role: "assistant", content: aiResponse }
            ];
          }
        } catch (error) {
          console.error("AI processing error:", error);
          resultInfo.textContent = "Error: " + (error.message || "Could not process data");
        }
      } else {
        resultInfo.textContent = "Please select a range with data first.";
      }
    });
  } catch (error) {
    console.error(error);
    const resultInfo = document.getElementById("result-info");
    resultInfo.textContent = "Error: " + (error.message || "An unknown error occurred");
  }
}

async function sendToOpenAI(prompt: string, isFollowUp: boolean = false, history = null): Promise<string> {
  try {
    // API key
    const apiKey = "YOUR_OPENAI_API_KEY"; // Replace with your actual API key
    
    // Prepare messages based on conversation history or just the current prompt
    let messages;
    
    if (isFollowUp && history && history.length > 0) {
      // Use provided history for follow-up questions
      messages = history;
    } else if (isFollowUp && conversationHistory.length > 0) {
      // Use conversation history for follow-up questions
      messages = conversationHistory;
    } else {
      // Just use the current prompt
      messages = [{ role: "user", content: prompt }];
    }
    
    const response = await fetch("https://api.openai.com/v1/chat/completions", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "Authorization": `Bearer ${apiKey}`
      },
      body: JSON.stringify({
        model: "gpt-4o-mini", // Using a more widely available model
        messages: messages
      })
    });
    
    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`API error (${response.status}): ${errorText}`);
    }
  
    const data = await response.json();
    
    // Check if the expected data structure exists
    if (data && data.choices && data.choices.length > 0 && data.choices[0].message) {
      return data.choices[0].message.content;
    } else {
      console.error("Unexpected API response format:", data);
      return "Error: Received unexpected response format from AI service.";
    }
  } catch (error) {
    console.error("Error calling OpenAI:", error);
    throw error; // Propagate the error to handle it in the main function
  }
}
