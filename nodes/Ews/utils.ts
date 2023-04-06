export const safeGetter = (getter: () => any) => {
	try {
		return getter()
	} catch (error) {
		return undefined
	}
}
