export interface IGenerator {
  create(): IGenerator;
  generateSlides(): Promise<void>;
  cleanup(): Promise<void>;
}
