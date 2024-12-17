/* global clearInterval, console, setInterval */
const { AzureOpenAI } = require("openai");

window.endpoint = "endpoint";
window.apiversion = "apiversion";
window.deployment = "deployment";
window.apikey = "apikey";




/**
  * AI.
  * @customfunction 
  * @param {string} msg AI.
  * @return {string}
  */
async function AI(msg) {
  try {
    //const client = new AzureOpenAI({ endpoint, apiKey, dangerouslyAllowBrowser: true, apiVersion, deployment });
    
    
    const endpoint = window.endpoint;
    const apiKey = window.apikey;
    const apiVersion = window.apiversion;
    const deployment = window.deployment;
    
    /*
    const endpoint = document.getElementById('endpoint').value;
    const apiKey = document.getElementById('apikey').value;
    const apiVersion = document.getElementById('apiversion').value;
    const deployment = document.getElementById('model').value;
    */

    const client = new AzureOpenAI({ endpoint, apiKey, dangerouslyAllowBrowser: true, apiVersion, deployment });
    const result = await client.chat.completions.create({
      messages: [
      { role: "system", content: "You are a helpful assistant." },
      { role: "user", content: msg },
      ],
      model: "",
    });
    /*
    for (const choice of result.choices) {
      return result.choices.map(choice => choice.message.content);
    }
    */
    return result.choices[0].message.content;
    //return "errorじゃありませｎ";
    //return window.endpoint + window.apikey + window.apiversion + window.deployment;
  }
  catch (error) {
    //return error;
    //return "errorです"+endpoint+apiKey;
    return error.toString();
  }
}



/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
export function add(first, second) {
  return first + second;
}

/**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
export function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
export function currentTime() {
  return new Date().toLocaleTimeString();
}

/**
 * Increments a value once a second.
 * @customfunction
 * @param {number} incrementBy Amount to increment
 * @param {CustomFunctions.StreamingInvocation<number>} invocation
 */
export function increment(incrementBy, invocation) {
  let result = 0;
  const timer = setInterval(() => {
    result += incrementBy;
    invocation.setResult(result);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}

/**
 * Writes a message to console.log().
 * @customfunction LOG
 * @param {string} message String to write.
 * @returns String to write.
 */
export function logMessage(message) {
  console.log(message);
  
  return message;
}

/**
 * Return
 * @customfunction GETVALUE
 * @returns {string} 
 */
export function getValue() {
  try {
    return window.endpoint + window.apiversion + window.deployment;
  }
  catch (error) {
    return error.toString();
  }
}