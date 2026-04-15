'use server';
/**
 * @fileOverview An AI assistant flow for suggesting or explaining Excel formulas.
 *
 * - aiAssistFormulaCreation - A function that suggests Excel formulas based on natural language or explains existing formulas.
 * - AIAssistFormulaCreationInput - The input type for the aiAssistFormulaCreation function.
 * - AIAssistFormulaCreationOutput - The return type for the aiAssistFormulaCreation function.
 */

import { ai } from '@/ai/genkit';
import { z } from 'genkit';

const AIAssistFormulaCreationInputSchema = z
  .object({
    description: z
      .string()
      .optional()
      .describe(
        'A natural language description for which an Excel formula needs to be suggested. Example: "calculate the total price by multiplying quantity and unit price"'
      ),
    formulaToExplain: z
      .string()
      .optional()
      .describe('An existing Excel formula that needs to be explained. Example: "=SUM(A2:A10)"'),
  })
  .refine(
    data =>
      (data.description !== undefined && data.formulaToExplain === undefined) ||
      (data.description === undefined && data.formulaToExplain !== undefined),
    'Either "description" or "formulaToExplain" must be provided, but not both.'
  );
export type AIAssistFormulaCreationInput = z.infer<
  typeof AIAssistFormulaCreationInputSchema
>;

const AIAssistFormulaCreationOutputSchema = z.object({
  suggestedFormula: z
    .string()
    .optional()
    .describe(
      'The Excel formula suggested by the AI based on the description provided. This field is present if a description was given.'
    ),
  explanation: z
    .string()
    .optional()
    .describe(
      'A clear and concise explanation of the provided Excel formula. This field is present if a formulaToExplain was given.'
    ),
});
export type AIAssistFormulaCreationOutput = z.infer<
  typeof AIAssistFormulaCreationOutputSchema
>;

export async function aiAssistFormulaCreation(
  input: AIAssistFormulaCreationInput
): Promise<AIAssistFormulaCreationOutput> {
  return aiAssistFormulaCreationFlow(input);
}

const formulaAssistantPrompt = ai.definePrompt({
  name: 'formulaAssistantPrompt',
  input: { schema: AIAssistFormulaCreationInputSchema },
  output: { schema: AIAssistFormulaCreationOutputSchema },
  prompt: `You are an expert Excel formula assistant. Your task is to either suggest an Excel formula based on a natural language description or explain an existing Excel formula.

If a description is provided, your primary goal is to provide a concise and correct Excel formula. If an existing formula is provided, your primary goal is to explain it clearly. Do not provide both a suggested formula and an explanation unless the description explicitly asks for both, in which case prioritize the request within the description. Assume standard Excel cell references (e.g., A2, B2) for suggestions.

{{#if description}}
Description for formula suggestion: {{{description}}}
{{else if formulaToExplain}}
Excel formula to explain: {{{formulaToExplain}}}
{{/if}}
`,
});

const aiAssistFormulaCreationFlow = ai.defineFlow(
  {
    name: 'aiAssistFormulaCreationFlow',
    inputSchema: AIAssistFormulaCreationInputSchema,
    outputSchema: AIAssistFormulaCreationOutputSchema,
  },
  async input => {
    const { output } = await formulaAssistantPrompt(input);
    return output!;
  }
);
